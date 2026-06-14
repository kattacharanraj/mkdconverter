/* =========================================================
   Theme switcher — cycles 3 themes and persists the choice.
   1) cyberpunk  2) portfolio (GitHub gold/black)  3) minimal (B&W)
   The initial theme is applied by an inline <head> script (no flash);
   this module just wires up the toggle button + label + meta color.
   ========================================================= */
(function () {
    'use strict';

    var THEMES = [
        { id: 'cyberpunk', label: 'Cyberpunk', color: '#00060E' },
        { id: 'portfolio', label: 'Black & Gold', color: '#050505' },
        { id: 'minimal',   label: 'Minimal',   color: '#ffffff' }
    ];

    var root = document.documentElement;
    var btn = document.getElementById('themeBtn');
    var label = document.getElementById('themeLabel');
    var meta = document.querySelector('meta[name="theme-color"]');

    function find(id) {
        for (var i = 0; i < THEMES.length; i++) {
            if (THEMES[i].id === id) return THEMES[i];
        }
        return THEMES[0];
    }

    function current() {
        return root.getAttribute('data-theme') || 'cyberpunk';
    }

    function apply(id) {
        var theme = find(id);
        root.setAttribute('data-theme', theme.id);
        try { localStorage.setItem('crk_theme', theme.id); } catch (e) { /* ignore */ }
        if (label) label.textContent = theme.label;
        if (meta) meta.setAttribute('content', theme.color);
    }

    // Sync UI to whatever the head script already applied.
    apply(current());

    if (btn) {
        btn.addEventListener('click', function () {
            var idx = 0;
            for (var i = 0; i < THEMES.length; i++) {
                if (THEMES[i].id === current()) { idx = i; break; }
            }
            apply(THEMES[(idx + 1) % THEMES.length].id);
        });
    }
})();
