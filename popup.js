function updateButtonStatus(button, success = true, message = '') {
  const defaultText = button.dataset.defaultText || button.textContent;
  const originalColor = button.style.backgroundColor;
  const statusText = message || (success ? '복사 완료!' : '복사 실패');

  button.textContent = statusText;
  button.style.backgroundColor = success ? '#2196F3' : '#f44336';

  setTimeout(() => {
    button.textContent = defaultText;
    button.style.backgroundColor = originalColor;
    button.dataset.defaultText = '';
  }, 3000);
}

// 줌 참가자 명단 복사 기능
// 핵심: 참가자 목록은 #webclient iframe 안에 있음
// allFrames 대신 메인 프레임에서 iframe.contentDocument로 접근
document
  .getElementById('copyZoomParticipants')
  .addEventListener('click', () => {
    const button = document.getElementById('copyZoomParticipants');
    button.dataset.defaultText = button.textContent;
    button.textContent = '초기화...';

    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      const tabId = tabs[0] && tabs[0].id;
      if (!tabId) {
        updateButtonStatus(button, false, '탭 없음');
        return;
      }

      const seenPositions = {};
      const allNames = [];
      let done = false;

      // 안전 타임아웃 60초
      const safetyTimer = setTimeout(() => {
        if (done) return;
        done = true;
        if (allNames.length > 0) {
          copyResult();
        } else {
          updateButtonStatus(button, false, '시간 초과');
        }
      }, 60000);

      // 메인 프레임에만 동기 함수 실행 (allFrames 사용 안 함)
      function execZoom(fn, callback) {
        chrome.scripting.executeScript(
          {
            target: { tabId },
            world: 'MAIN',
            function: fn,
          },
          (results) => {
            if (chrome.runtime.lastError) {
              callback({ error: chrome.runtime.lastError.message });
              return;
            }
            var r = results && results[0] && results[0].result;
            callback(r || null);
          }
        );
      }

      // 1단계: 초기화
      execZoom(zoomInit, function (initResult) {
        if (done) return;
        if (!initResult || initResult.error || initResult.skip) {
          done = true;
          clearTimeout(safetyTimer);
          var msg = '참가자 없음';
          if (initResult && initResult.error) msg = initResult.error;
          if (initResult && initResult.msg) msg = initResult.msg;
          updateButtonStatus(button, false, msg);
          return;
        }

        button.textContent = '수집 중...';
        setTimeout(doCollectStep, 300);
      });

      // 2단계: 반복 수집
      function doCollectStep() {
        if (done) return;
        execZoom(zoomCollectAndScroll, function (stepResult) {
          if (done) return;
          if (!stepResult || stepResult.skip) {
            finishCollection();
            return;
          }

          (stepResult.items || []).forEach(function (item) {
            if (!seenPositions[item.t]) {
              seenPositions[item.t] = true;
              allNames.push(item.n);
            }
          });

          button.textContent = '수집 중... ' + allNames.length + '명';

          if (stepResult.atBottom) {
            finishCollection();
          } else {
            setTimeout(doCollectStep, 250);
          }
        });
      }

      // 3단계: 마무리
      function finishCollection() {
        if (done) return;
        done = true;
        clearTimeout(safetyTimer);

        execZoom(zoomReset, function () {
          copyResult();
        });
      }

      function copyResult() {
        if (allNames.length > 0) {
          navigator.clipboard
            .writeText(allNames.join('\n'))
            .then(() =>
              updateButtonStatus(button, true, allNames.length + '명 복사!')
            )
            .catch(() =>
              updateButtonStatus(button, false, '클립보드 실패')
            );
        } else {
          updateButtonStatus(button, false, '참가자 없음');
        }
      }
    });
  });

// ====== 줌 동기 함수들 ======
// #webclient iframe 내부의 contentDocument에서 참가자 DOM 접근
// allFrames 사용하지 않고, 메인 프레임에서 iframe 직접 접근

// iframe 안의 참가자 문서(doc) 가져오기 헬퍼 (각 함수 내부에 인라인)
// - 먼저 현재 document에서 찾고
// - 없으면 #webclient iframe의 contentDocument에서 찾음

function zoomInit() {
  try {
    // 참가자 컨테이너 찾기: 현재 문서 → iframe 순서
    var doc = document;
    var c = doc.querySelector('.participants-list-container');
    if (!c) {
      var iframe = document.getElementById('webclient');
      if (iframe && iframe.contentDocument) {
        doc = iframe.contentDocument;
        c = doc.querySelector('.participants-list-container');
      }
    }
    if (!c) return { skip: true, msg: 'no-container' };

    c.scrollTop = 0;
    c.dispatchEvent(new Event('scroll', { bubbles: true }));
    return { ok: true };
  } catch (e) {
    return { skip: true, msg: 'init-err:' + e.message };
  }
}

function zoomCollectAndScroll() {
  try {
    var doc = document;
    var c = doc.querySelector('.participants-list-container');
    if (!c) {
      var iframe = document.getElementById('webclient');
      if (iframe && iframe.contentDocument) {
        doc = iframe.contentDocument;
        c = doc.querySelector('.participants-list-container');
      }
    }
    if (!c) return { skip: true };

    // 현재 보이는 참가자 수집
    var items = [];
    var posEls = doc.querySelectorAll('.participants-item-position');
    posEls.forEach(function (posEl) {
      var top = posEl.style.top;
      if (!top) return;
      var nameEl = posEl.querySelector('.participants-item__display-name');
      if (nameEl) {
        var name = nameEl.textContent.trim();
        if (name) items.push({ t: top, n: name });
      }
    });

    // 스크롤 진행
    var maxScroll = c.scrollHeight - c.clientHeight;
    var atBottom = c.scrollTop >= maxScroll - 2;
    if (!atBottom) {
      var step = Math.floor(c.clientHeight * 0.5);
      c.scrollTop = Math.min(c.scrollTop + step, maxScroll);
      c.dispatchEvent(new Event('scroll', { bubbles: true }));
    }

    return { items: items, atBottom: atBottom };
  } catch (e) {
    return { skip: true };
  }
}

function zoomReset() {
  try {
    var doc = document;
    var c = doc.querySelector('.participants-list-container');
    if (!c) {
      var iframe = document.getElementById('webclient');
      if (iframe && iframe.contentDocument) {
        doc = iframe.contentDocument;
        c = doc.querySelector('.participants-list-container');
      }
    }
    if (c) {
      c.scrollTop = 0;
      c.dispatchEvent(new Event('scroll', { bubbles: true }));
    }
    return { ok: true };
  } catch (e) {
    return { ok: false };
  }
}

