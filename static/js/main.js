function toggleTheme() {
            const c = document.documentElement.getAttribute('data-theme') || 'dark';
            const n = c === 'dark' ? 'light' : 'dark';
            document.documentElement.setAttribute('data-theme', n);
            localStorage.setItem('theme', n);
            updateThemeBtnText(n);
        }
        function updateThemeBtnText(t) {
            const b = document.getElementById('theme-toggle-btn');
            if (b) b.innerHTML = t === 'dark' ? '☀️' : '🌙';
        }
        function openSidebar() {
            document.body.classList.remove('sidebar-collapsed');
            saveSidebarPref();
        }
        function closeSidebar() {
            document.body.classList.add('sidebar-collapsed');
            saveSidebarPref();
        }
        function toggleSidebar() {
            if (document.body.classList.contains('sidebar-collapsed')) openSidebar();
            else closeSidebar();
        }
        function initSidebarState() {
            const pref = localStorage.getItem('sidebar_pref');
            if (pref === 'collapsed') {
                document.body.classList.add('sidebar-collapsed');
                return;
            }
            if (pref === 'expanded') {
                document.body.classList.remove('sidebar-collapsed');
                return;
            }
            // Sin preferencia: mobile colapsa, desktop expande
            if (window.innerWidth <= 992) document.body.classList.add('sidebar-collapsed');
            else document.body.classList.remove('sidebar-collapsed');
        }

        function saveSidebarPref() {
            const state = document.body.classList.contains('sidebar-collapsed') ? 'collapsed' : 'expanded';
            localStorage.setItem('sidebar_pref', state);
        }
        let _lastWidth = window.innerWidth;
        window.addEventListener('resize', function() {
            const nowMobile = window.innerWidth <= 992;
            const wasMobile = _lastWidth <= 992;
            // Solo re-inicializar si cambia de mobile <-> desktop Y no hay preferencia manual
            if (nowMobile !== wasMobile && !localStorage.getItem('sidebar_pref')) {
                initSidebarState();
            }
            _lastWidth = window.innerWidth;
        });
        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape' && window.innerWidth <= 992) closeSidebar();
        });
        document.addEventListener('click', function(e) {
            if (window.innerWidth > 992) return;
            if (e.target.closest('.sidebar-overlay')) closeSidebar();
        });
        let _touchStartX = 0, _touchStartY = 0;
        document.addEventListener('touchstart', function(e) {
            if (window.innerWidth > 992) return;
            _touchStartX = e.touches[0].clientX;
            _touchStartY = e.touches[0].clientY;
        }, { passive: true });
        document.addEventListener('touchend', function(e) {
            if (window.innerWidth > 992) return;
            const dx = e.changedTouches[0].clientX - _touchStartX;
            const dy = Math.abs(e.changedTouches[0].clientY - _touchStartY);
            if (dy > 60) return;
            if (_touchStartX < 30 && dx > 60) openSidebar();
            if (!document.body.classList.contains('sidebar-collapsed') && dx < -60) closeSidebar();
        }, { passive: true });

        function initToasts() {
            const fl = document.querySelectorAll('.flash');
            if (fl.length === 0) return;
            const c = document.createElement('div');
            c.className = 'toast-container';
            fl[0].parentNode.insertBefore(c, fl[0]);
            fl.forEach(f => {
                c.appendChild(f);
                const dismiss = () => { f.classList.add('dismissing'); setTimeout(() => f.remove(), 320); };
                setTimeout(dismiss, 4500);
                f.addEventListener('click', dismiss);
            });
        }
        function initRipple() {
            document.addEventListener('click', function(e) {
                const b = e.target.closest('button, .btn');
                if (!b || b.disabled) return;
                if (b.classList.contains('btn-theme') || b.classList.contains('btn-logout') || b.classList.contains('btn-theme-nav') || b.classList.contains('navbar-brand')) return;
                const r = b.getBoundingClientRect();
                const size = Math.max(r.width, r.height);
                const x = e.clientX - r.left - size / 2;
                const y = e.clientY - r.top - size / 2;
                const rip = document.createElement('span');
                rip.className = 'ripple';
                rip.style.width = rip.style.height = size + 'px';
                rip.style.left = x + 'px';
                rip.style.top = y + 'px';
                b.appendChild(rip);
                setTimeout(() => rip.remove(), 700);
            });
        }
            function initCardEntrance() {
                if (!('IntersectionObserver' in window)) return;
                const allCards = document.querySelectorAll('.card, .stat-card');
                const cards = Array.from(allCards).slice(0, 20);
                cards.forEach((c, i) => {
                    c.style.opacity = '0';
                    c.style.transform = 'translateY(12px)';
                    c.style.transition = 'opacity 0.5s cubic-bezier(0.16,1,0.3,1), transform 0.5s cubic-bezier(0.16,1,0.3,1)';
                    c.style.transitionDelay = Math.min(i * 40, 250) + 'ms';
                });

                // ⚠️ CRÍTICO: al terminar la animación, quitar el transform por completo.
                // Si dejamos transform: translateY(0), la card crea un containing block
                // y los hijos con position: fixed se posicionan respecto a la card (no al viewport),
                // además de quedar recortados por overflow: hidden.
                const cleanup = (el) => {
                    el.style.transform = '';
                    el.style.transition = '';
                    el.style.transitionDelay = '';
                };

                const obs = new IntersectionObserver(entries => {
                    entries.forEach(en => {
                        if (en.isIntersecting) {
                            en.target.style.opacity = '1';
                            en.target.style.transform = 'translateY(0)';
                            // Espera a que termine la transición (0.5s + hasta 250ms de delay)
                            setTimeout(() => cleanup(en.target), 900);
                            obs.unobserve(en.target);
                        }
                    });
                }, { threshold: 0.05 });

                cards.forEach(c => obs.observe(c));

                // Fallback para cards que nunca lleguen a intersectar
                setTimeout(() => {
                    cards.forEach(c => {
                        if (c.style.opacity === '0') {
                            c.style.opacity = '1';
                            cleanup(c);
                        }
                    });
                }, 1500);
            }

window.toggleQuickCalculator = function(force) {
    const p = document.getElementById('quickCalcPanel');
    if (!p) return;
    const isOpen = p.classList.contains('show');
    const shouldOpen = (force === undefined) ? !isOpen : !!force;

    if (shouldOpen) {
        p.style.display = 'flex';
        // Pequeño delay para la animación CSS
        setTimeout(() => p.classList.add('show'), 10);
        
        // Actualizar tasas al abrir
        if (typeof qcRefreshRates === 'function') qcRefreshRates();
        
        // Enfocar el input del conversor por defecto
        setTimeout(() => {
            const amt = document.getElementById('qcFxAmount');
            if(amt) amt.focus();
        }, 150);
    } else {
        p.classList.remove('show');
        setTimeout(() => p.style.display = 'none', 300);
    }
};

function qcUpdateExpression() {
    const d = document.getElementById('qcDisplay');
    const ex = document.getElementById('qcExpression');
    if (ex && d) ex.textContent = d.value === '0' ? '' : d.value;
}

function qcAppend(v) {
    const d = document.getElementById('qcDisplay');
    if (!d) return;
    if (d.value === '0' && /[0-9.]/.test(v)) d.value = v;
    else if (d.value === 'Error' || d.value === '0') d.value = v;
    else d.value += v;
    qcUpdateExpression();
    if (navigator.vibrate) navigator.vibrate(10);
}

function qcClear() {
    const d = document.getElementById('qcDisplay');
    if (d) d.value = '0';
    const ex = document.getElementById('qcExpression');
    if(ex) ex.textContent = '';
    if (navigator.vibrate) navigator.vibrate(15);
}

function qcBackspace() {
    const d = document.getElementById('qcDisplay');
    if (!d) return;
    d.value = d.value.length > 1 ? d.value.slice(0, -1) : '0';
    qcUpdateExpression();
    if (navigator.vibrate) navigator.vibrate(10);
}

function qcCalculate() {
    const d = document.getElementById('qcDisplay');
    const ex = document.getElementById('qcExpression');
    if (!d) return;
    try {
        if (!/^[0-9+\-*/().\s]+$/.test(d.value) || d.value.trim() === '') throw 0;
        const result = Function('"use strict";return (' + d.value + ')')();
        if (!Number.isFinite(result)) throw 0;
        
        if (ex) ex.textContent = d.value + ' =';
        d.value = String(Math.round((result + Number.EPSILON) * 1e10) / 1e10);
        if (navigator.vibrate) navigator.vibrate([20, 30, 20]);
    } catch (e) {
        d.value = 'Error';
        if (ex) ex.textContent = '';
        if (navigator.vibrate) navigator.vibrate([50, 50, 50]);
    }
}

function qcGetRates() {
    const usdEl = document.getElementById('rate_usd_display');
    const eurEl = document.getElementById('rate_eur_display');
    const usd = parseFloat((usdEl && usdEl.textContent) || '0') || 0;
    const eur = parseFloat((eurEl && eurEl.textContent) || '0') || 0;
    return { usd, eur };
}

function qcToBs(amount, currency, r) {
    if (currency === 'VES') return amount;
    if (currency === 'USD') return amount * r.usd;
    if (currency === 'EUR') return amount * r.eur;
    return amount;
}

function qcFromBs(amountBs, currency, r) {
    if (currency === 'VES') return amountBs;
    if (currency === 'USD') return r.usd > 0 ? amountBs / r.usd : 0;
    if (currency === 'EUR') return r.eur > 0 ? amountBs / r.eur : 0;
    return amountBs;
}

function qcFmt(n, decimals) {
    if (!Number.isFinite(n)) return '0';
    decimals = (decimals === undefined) ? 2 : decimals;
    return n.toLocaleString('es-VE', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

function qcConvertFx() {
    const amountEl = document.getElementById('qcFxAmount');
    const fromEl   = document.getElementById('qcFxFrom');
    const toEl     = document.getElementById('qcFxTo');
    const resEl    = document.getElementById('qcFxResult');
    const ratesEl  = document.getElementById('qcFxRates');
    if (!amountEl || !fromEl || !toEl || !resEl) return;

    const amount = parseFloat(amountEl.value) || 0;
    const from = fromEl.value;
    const to   = toEl.value;
    const r    = qcGetRates();

    const bs = qcToBs(amount, from, r);
    const result = qcFromBs(bs, to, r);
    const symbols = { USD: '$ ', EUR: '€ ', VES: 'Bs ' };
    
    resEl.textContent = (symbols[to] || '') + qcFmt(result, 2);
    resEl.dataset.rawResult = (symbols[to] || '') + qcFmt(result, 2);

    if (ratesEl) {
        if (r.usd <= 0 && r.eur <= 0) {
            ratesEl.innerHTML = '<span style="color:var(--danger);">⚠️ No hay tasas configuradas.</span>';
            return;
        }
        let html = '';
        html += '💵 <b>1 USD</b> = <span style="color:var(--accent);font-weight:800">' + qcFmt(r.usd, 2) + ' Bs</span> &nbsp;|&nbsp; ';
        html += '💶 <b>1 EUR</b> = <span style="color:var(--accent);font-weight:800">' + qcFmt(r.eur, 2) + ' Bs</span>';
        ratesEl.innerHTML = html;
    }
}

function qcSwapFx() {
    const fromEl = document.getElementById('qcFxFrom');
    const toEl   = document.getElementById('qcFxTo');
    if (!fromEl || !toEl) return;
    const tmp = fromEl.value;
    fromEl.value = toEl.value;
    toEl.value = tmp;
    qcConvertFx();
    if (navigator.vibrate) navigator.vibrate(15);
}

function qcQuickFx(amount) {
    const a = document.getElementById('qcFxAmount');
    const f = document.getElementById('qcFxFrom');
    const t = document.getElementById('qcFxTo');
    if (a) a.value = amount;
    if (f) f.value = 'USD';
    if (t) t.value = 'VES';
    qcConvertFx();
    if (navigator.vibrate) navigator.vibrate(10);
}

function qcCopyResult() {
    const resEl = document.getElementById('qcFxResult');
    const toast = document.getElementById('qcCopyToast');
    if (!resEl) return;
    const textToCopy = resEl.dataset.rawResult || resEl.textContent;
    
    navigator.clipboard.writeText(textToCopy).then(() => {
        if (toast) {
            toast.classList.add('show');
            setTimeout(() => toast.classList.remove('show'), 2000);
        }
    }).catch(() => {});
}

function qcRefreshRates() {
    if (typeof pollExchangeRates === 'function') pollExchangeRates();
    setTimeout(qcConvertFx, 350);
    qcConvertFx();
}            

        function pollExchangeRates() {
            fetch("/api/rates/check", { credentials: 'same-origin' })
                .then(r => r.ok ? r.json() : null)
                .then(d => {
                    if (!d) return;
                    const u = document.getElementById('rate_usd_display');
                    const e = document.getElementById('rate_eur_display');
                    if (u) u.textContent = parseFloat(d.tasa_usd).toFixed(2);
                    if (e) e.textContent = parseFloat(d.tasa_eur).toFixed(2);
                }).catch(() => {});
        }


 (function() {
    'use strict';

    function savePos(panelEl, key) {
        try {
            const rect = panelEl.getBoundingClientRect();
            localStorage.setItem(key, JSON.stringify({
                left: Math.round(rect.left),
                top: Math.round(rect.top)
            }));
        } catch(e) {}
    }

    function restorePos(panelEl, key) {
        try {
            const saved = JSON.parse(localStorage.getItem(key) || 'null');
            if (!saved || typeof saved.left !== 'number' || typeof saved.top !== 'number') return;
            const vw = window.innerWidth;
            const vh = window.innerHeight;
            const panelW = panelEl.offsetWidth || 380;
            const panelH = panelEl.offsetHeight || 400;
            const left = Math.max(4, Math.min(vw - panelW - 4, saved.left));
            const top  = Math.max(4, Math.min(vh - panelH - 4, saved.top));
            panelEl.style.left = left + 'px';
            panelEl.style.top = top + 'px';
            panelEl.style.right = 'auto';
            panelEl.style.bottom = 'auto';
        } catch(e) {}
    }

    function makeDraggable(panelEl, storageKey) {
        if (!panelEl) return;
        const handle = panelEl.querySelector('[data-drag-handle]');
        if (!handle) return;

        let dragging = false;
        let moved = false;
        let startX = 0, startY = 0;
        let startLeft = 0, startTop = 0;
        let pointerId = null;

        handle.addEventListener('pointerdown', function(e) {
            // Ignorar clicks sobre elementos interactivos dentro del handle
            if (e.target.closest('button, input, select, a, textarea')) return;
            e.preventDefault();
            dragging = true;
            moved = false;
            pointerId = e.pointerId;

            const rect = panelEl.getBoundingClientRect();
            startX = e.clientX;
            startY = e.clientY;
            startLeft = rect.left;
            startTop = rect.top;

            handle.classList.add('dragging');
            try { handle.setPointerCapture(e.pointerId); } catch(_) {}
        });

        handle.addEventListener('pointermove', function(e) {
            if (!dragging) return;
            const dx = e.clientX - startX;
            const dy = e.clientY - startY;
            if (Math.abs(dx) > 3 || Math.abs(dy) > 3) moved = true;

            const vw = window.innerWidth;
            const vh = window.innerHeight;
            const panelW = panelEl.offsetWidth;
            const panelH = panelEl.offsetHeight;

            let newLeft = startLeft + dx;
            let newTop  = startTop + dy;
            newLeft = Math.max(4, Math.min(vw - panelW - 4, newLeft));
            newTop  = Math.max(4, Math.min(vh - panelH - 4, newTop));

            panelEl.style.left = newLeft + 'px';
            panelEl.style.top = newTop + 'px';
            panelEl.style.right = 'auto';
            panelEl.style.bottom = 'auto';
        });

        function stopDrag() {
            if (!dragging) return;
            dragging = false;
            handle.classList.remove('dragging');
            try {
                if (pointerId !== null) handle.releasePointerCapture(pointerId);
            } catch(_) {}
            pointerId = null;
            if (moved) savePos(panelEl, storageKey);
        }

        handle.addEventListener('pointerup', stopDrag);
        handle.addEventListener('pointercancel', stopDrag);
        handle.addEventListener('lostpointercapture', stopDrag);

        // Restaurar al cargar
        restorePos(panelEl, storageKey);

        // Reajustar al redimensionar ventana
        window.addEventListener('resize', function() {
            const rect = panelEl.getBoundingClientRect();
            const vw = window.innerWidth;
            const vh = window.innerHeight;
            if (rect.right > vw || rect.bottom > vh) {
                restorePos(panelEl, storageKey);
            }
        });
    }

    window.rsvMakeDraggable = makeDraggable;

    document.addEventListener('DOMContentLoaded', function() {
        // Limpiar posiciones viejas
        localStorage.removeItem('rsv_calc_pos_v1');
        localStorage.removeItem('rsv_search_pos_v1');
        
        makeDraggable(document.getElementById('quickCalcPanel'), 'rsv_calc_pos_v1');
        // Se elimina el makeDraggable para globalSearchPanel para evitar que se mueva de forma extraña
        
        // Atajos de teclado
        document.addEventListener('keydown', function(e) {
            // Ctrl+K o Cmd+K para Lupa
            if ((e.ctrlKey || e.metaKey) && (e.key === 'k' || e.key === 'K')) {
                e.preventDefault();
                toggleGlobalSearch();
            }
            // Ctrl+J o Cmd+J para Calculadora
            if ((e.ctrlKey || e.metaKey) && (e.key === 'j' || e.key === 'J')) {
                e.preventDefault();
                toggleQuickCalculator();
            }
            // Escape para cerrar ambos
            if (e.key === 'Escape') {
                const p1 = document.getElementById('quickCalcPanel');
                if (p1 && p1.classList.contains('show')) toggleQuickCalculator(false);
                const p2 = document.getElementById('globalSearchPanel');
                if (p2 && p2.classList.contains('open')) toggleGlobalSearch(false);
            }
        });
    });

})();       