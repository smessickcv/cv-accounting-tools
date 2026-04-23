/**
 * CV Internal Tools — version-check.js
 *
 * Each tool page must define (before this script loads):
 *   const TOOL_KEY     = 'some-key';   // matches a key in version.json
 *   const TOOL_VERSION = '1.0';        // current hardcoded version
 *
 * On load this script fetches /version.json and, if a newer version exists,
 * injects a sticky banner prompting the user to get the update from #internal-tools.
 */
(async function cvVersionCheck() {
  if (typeof TOOL_KEY === 'undefined' || typeof TOOL_VERSION === 'undefined') return;

  try {
    const origin = location.origin === 'null' ? '' : location.origin;
    const base   = origin + (location.pathname.replace(/\/[^/]*$/, '') || '');
    const url    = base + '/version.json?t=' + Date.now();

    const resp = await fetch(url);
    if (!resp.ok) return;

    const data   = await resp.json();
    const latest = data && data.tools && data.tools[TOOL_KEY];
    if (!latest) return;

    if (isNewer(latest, String(TOOL_VERSION))) {
      showBanner(latest);
    }
  } catch (_) { /* offline or opened as local file — silently skip */ }

  function isNewer(a, b) {
    const pa = a.split('.').map(Number), pb = b.split('.').map(Number);
    for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
      const d = (pa[i] || 0) - (pb[i] || 0);
      if (d > 0) return true;
      if (d < 0) return false;
    }
    return false;
  }

  function showBanner(latest) {
    if (document.getElementById('cv-update-banner')) return; // already shown

    const banner = document.createElement('div');
    banner.id = 'cv-update-banner';
    Object.assign(banner.style, {
      background:  '#b91c1c',
      color:       '#fff',
      padding:     '10px 20px',
      fontSize:    '13px',
      fontFamily:  '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
      textAlign:   'center',
      lineHeight:  '1.5',
      position:    'sticky',
      top:         '0',
      zIndex:      '99999',
      display:     'flex',
      alignItems:  'center',
      justifyContent: 'center',
      gap:         '12px',
    });

    banner.innerHTML =
      '<span>&#9650; <strong>Update available: v' + latest + '</strong> &mdash; ' +
      'Get the latest version from <strong>#internal-tools</strong> on Slack.</span>' +
      '<button onclick="document.getElementById(\'cv-update-banner\').remove()" ' +
      'style="background:transparent;border:1px solid rgba(255,255,255,0.6);color:#fff;' +
      'padding:3px 10px;border-radius:3px;cursor:pointer;font-size:12px;flex-shrink:0">Dismiss</button>';

    // Insert at very top of body (before any existing children)
    if (document.body) {
      document.body.insertBefore(banner, document.body.firstChild);
    } else {
      document.addEventListener('DOMContentLoaded', () =>
        document.body.insertBefore(banner, document.body.firstChild));
    }
  }
})();
