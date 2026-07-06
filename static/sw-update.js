(function () {
  if (!('serviceWorker' in navigator)) return;

  const APP_VERSION = document.documentElement.dataset.appVersion || '';
  let reloaded = false;
  let waitingWorker = null;

  function showUpdateToast(worker, source) {
    waitingWorker = worker || waitingWorker;
    if (document.getElementById('sw-update-toast')) return;

    const toast = document.createElement('div');
    toast.id = 'sw-update-toast';
    toast.setAttribute('role', 'alert');
    toast.innerHTML = `
      <div class="sw-update-copy">
        <strong>已有新版可使用</strong>
        <span>重新整理後會載入最新的考卷格式校正系統。</span>
      </div>
      <div class="sw-update-actions">
        <button id="sw-update-now" type="button">重新整理</button>
        <button id="sw-update-dismiss" type="button" aria-label="稍後再說">×</button>
      </div>
    `;
    document.body.appendChild(toast);

    document.getElementById('sw-update-now').addEventListener('click', () => {
      toast.remove();
      if (waitingWorker) {
        waitingWorker.postMessage({ type: 'SKIP_WAITING' });
      } else {
        window.location.reload();
      }
    });
    document.getElementById('sw-update-dismiss').addEventListener('click', () => toast.remove());
    console.info('[SW] update available', source || '');
  }

  function watchWorker(worker) {
    if (!worker) return;
    worker.addEventListener('statechange', () => {
      if (worker.state === 'installed' && navigator.serviceWorker.controller) {
        showUpdateToast(worker, 'installed');
      }
    });
  }

  async function checkVersion() {
    if (!APP_VERSION) return;
    try {
      const response = await fetch(`version.json?t=${Date.now()}`, { cache: 'no-store' });
      if (!response.ok) return;
      const data = await response.json();
      if (data.version && data.version !== APP_VERSION) {
        showUpdateToast(waitingWorker, `version:${data.version}`);
      }
    } catch (error) {
      console.debug('[SW] version check skipped', error);
    }
  }

  window.addEventListener('load', () => {
    navigator.serviceWorker
      .register('sw.js', { updateViaCache: 'none' })
      .then((registration) => {
        if (registration.waiting && navigator.serviceWorker.controller) {
          showUpdateToast(registration.waiting, 'waiting');
        }
        watchWorker(registration.installing);
        registration.addEventListener('updatefound', () => watchWorker(registration.installing));

        window.addEventListener('focus', checkVersion);
        window.addEventListener('online', checkVersion);
        window.addEventListener('pageshow', checkVersion);
        document.addEventListener('visibilitychange', () => {
          if (document.visibilityState === 'visible') checkVersion();
        });

        setTimeout(checkVersion, 8000);
        setInterval(() => {
          registration.update().catch(() => null);
          checkVersion();
        }, 180000);
      })
      .catch((error) => console.debug('[SW] registration skipped', error));
  });

  navigator.serviceWorker.addEventListener('message', (event) => {
    if (event.data && event.data.type === 'SW_ACTIVATED' && event.data.version !== APP_VERSION) {
      showUpdateToast(null, `activated:${event.data.version}`);
    }
  });

  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (reloaded) return;
    reloaded = true;
    window.location.reload();
  });
})();
