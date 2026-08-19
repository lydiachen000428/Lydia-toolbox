(function () {
  // Load Tabler icons if not already present
  if (!document.querySelector('link[href*="tabler-icons"]')) {
    var iconLink = document.createElement('link');
    iconLink.rel = 'stylesheet';
    iconLink.href = 'https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2.44.0/tabler-icons.min.css';
    document.head.appendChild(iconLink);
  }

  var NAV_W = 148;

  var items = [
    { icon: 'ti-home',            label: '首頁',       path: '/' },
    { icon: 'ti-package',         label: 'Packing List', path: '/packing' },
    { icon: 'ti-database',        label: '料號資料庫', path: '/items' },
    { icon: 'ti-filter',          label: '訂單篩選',   path: '/tools/orders' },
    { icon: 'ti-list-check',      label: 'PO出貨文字化', path: '/tools/po-formatter' },
    { icon: 'ti-shield-check',    label: 'CBP 助手',   path: '/tools/cbp' },
    { icon: 'ti-currency-dollar', label: '匯率換算',   path: '/tools/currency' },
  ];

  var cur = window.location.pathname;
  // For root, only match exactly
  function isActive(path) {
    if (path === '/') return cur === '/';
    return cur === path || cur.startsWith(path + '/');
  }

  // Inject sidebar CSS
  var style = document.createElement('style');
  style.textContent = [
    '#_nav{position:fixed;left:0;top:0;bottom:0;width:' + NAV_W + 'px;',
    'background:#f5f4f0;border-right:0.5px solid rgba(0,0,0,0.1);',
    'display:flex;flex-direction:column;z-index:999;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Noto Sans TC",sans-serif;}',
    '#_nav-logo{font-size:10px;font-weight:500;letter-spacing:0.1em;color:#aaa;',
    'padding:18px 14px 12px;text-transform:uppercase;}',
    '#_nav a{display:flex;align-items:center;gap:8px;padding:8px 14px;',
    'font-size:12px;color:#666;text-decoration:none;border-right:2px solid transparent;',
    'transition:background 0.1s,color 0.1s;}',
    '#_nav a:hover{background:rgba(0,0,0,0.04);color:#1a1a1a;}',
    '#_nav a.on{background:#fff;color:#1a1a1a;font-weight:500;border-right-color:#c17f3b;}',
    '#_nav a i{font-size:16px;flex-shrink:0;}',
    '#_nav a span{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}',
    'body{margin-left:' + NAV_W + 'px !important;}',
  ].join('');
  document.head.appendChild(style);

  // Build sidebar
  var nav = document.createElement('nav');
  nav.id = '_nav';

  var logo = document.createElement('div');
  logo.id = '_nav-logo';
  logo.textContent = "Lydia's Toolbox";
  nav.appendChild(logo);

  items.forEach(function (item) {
    var a = document.createElement('a');
    a.href = item.path;
    if (isActive(item.path)) a.className = 'on';
    a.innerHTML = '<i class="ti ' + item.icon + '" aria-hidden="true"></i><span>' + item.label + '</span>';
    nav.appendChild(a);
  });

  document.body.insertBefore(nav, document.body.firstChild);
})();
