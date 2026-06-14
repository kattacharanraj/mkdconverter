/* =========================================================
   Spider-Verse glitch effects (vanilla port)
   - Buttons: cyan/magenta TEXT tear layers (button stays still)
   - Hero title: scramble-on-hover with yellow reveal block
   ========================================================= */
(function () {
    'use strict';

    var letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

    function makeEl(tag, cls, text) {
        var el = document.createElement(tag);
        if (cls) el.className = cls;
        if (text != null) el.textContent = text;
        el.setAttribute('aria-hidden', 'true');
        return el;
    }

    /* ---- Inject glitch layers into every button ---- */
    function initButtons() {
        var buttons = document.querySelectorAll('.btn');
        for (var i = 0; i < buttons.length; i++) {
            var btn = buttons[i];
            if (btn.dataset.glitchInit) continue;
            btn.dataset.glitchInit = '1';

            var label = (btn.textContent || '').trim();
            if (!label) continue;

            var layers = makeEl('span', 'glitch-layers');
            layers.appendChild(makeEl('span', 'glitch-layer glitch-layer-1', label));
            layers.appendChild(makeEl('span', 'glitch-layer glitch-layer-2', label));

            btn.appendChild(layers);
        }
    }

    /* ---- Scramble + yellow reveal on the hero title ---- */
    function initTitle() {
        var h1 = document.querySelector('.hero h1.glitch-title');
        if (!h1 || h1.dataset.glitchInit) return;
        h1.dataset.glitchInit = '1';

        var fullText = (h1.getAttribute('data-text') || h1.textContent || '').trim();
        var reveal = makeEl('span', 'glitch-reveal', fullText);
        h1.appendChild(reveal);

        var iv = null;

        h1.addEventListener('mouseenter', function () {
            var iteration = 0;
            if (iv) clearInterval(iv);
            iv = setInterval(function () {
                reveal.textContent = fullText.split('').map(function (ch, idx) {
                    if (ch === ' ') return ' ';
                    if (idx < iteration) return fullText[idx];
                    return letters[Math.floor(Math.random() * letters.length)];
                }).join('');
                if (iteration >= fullText.length) { clearInterval(iv); iv = null; }
                iteration += 1 / 3;
            }, 30);
            h1.classList.add('revealing');
        });

        h1.addEventListener('mouseleave', function () {
            if (iv) { clearInterval(iv); iv = null; }
            reveal.textContent = fullText;
            h1.classList.remove('revealing');
        });
    }

    function init() {
        initButtons();
        initTitle();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
