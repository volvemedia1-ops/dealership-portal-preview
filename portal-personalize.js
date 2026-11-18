/**
 * Credit Builder Flow • Shared Portal Personalization & Link Forwarding
 * Supports Option A1 URL parameters, reset=1, localStorage fallback,
 * neutral defaults (no dealership/contact leakage), and #hash-safe link propagation.
 */
(function () {
    'use strict';

    var KEYS = ['dealer', 'contact', 'dp', 'date', 'declines', 'margin', 'banks', 'concern'];
    var VALID_CONCERNS = {
        exclusivity: true,
        popia: true,
        staff: true,
        cost: true,
        commitment: true
    };
    var STALE_DEFAULTS = {
        'Volkswagen Auto Group': true,
        'Johan Van Der Merwe': true,
        'Johan Botha': true,
        'Apex Auto Group': true
    };

    function sanitizeParam(val) {
        if (val === null || val === undefined) return '';
        var str = String(val).replace(/\+/g, ' ').trim();
        if (STALE_DEFAULTS[str]) return '';
        return str;
    }

    function safeGetStorage(key) {
        try {
            var val = localStorage.getItem('cbf_' + key);
            return sanitizeParam(val);
        } catch (e) {
            return '';
        }
    }

    function safeSetStorage(key, val) {
        try {
            if (val) {
                localStorage.setItem('cbf_' + key, val);
            } else {
                localStorage.removeItem('cbf_' + key);
            }
        } catch (e) {
            // Ignore storage errors in restricted contexts
        }
    }

    function clearAllStorage() {
        try {
            for (var i = 0; i < KEYS.length; i++) {
                localStorage.removeItem('cbf_' + KEYS[i]);
            }
        } catch (e) {
            // Ignore storage errors
        }
    }

    function toDealerSlug(name) {
        var slug = String(name || 'your-dealership')
            .toLowerCase()
            .replace(/[^a-z0-9]+/g, '-')
            .replace(/^-+|-+$/g, '');
        return slug || 'your-dealership';
    }

    function parseProspect() {
        var params = new URLSearchParams(window.location.search);
        var resetVal = (params.get('reset') || '').toLowerCase();
        if (resetVal === '1' || resetVal === 'true') {
            clearAllStorage();
        }

        var urlValues = {};
        var hasUrlParams = false;
        for (var i = 0; i < KEYS.length; i++) {
            var k = KEYS[i];
            if (params.has(k)) {
                hasUrlParams = true;
                urlValues[k] = sanitizeParam(params.get(k));
            }
        }

        // Persist ONLY when URL params are present.
        // When a fresh dealer entry URL arrives, clear stale optional keys not present in the URL.
        if (hasUrlParams) {
            var prevSavedDealer = safeGetStorage('dealer');
            var isEntryPage = /(?:^|\/)(?:index\.html)?$/i.test(window.location.pathname);
            var switchingDealer = urlValues.dealer && prevSavedDealer && urlValues.dealer !== prevSavedDealer;
            if (switchingDealer || (isEntryPage && urlValues.dealer)) {
                for (var j = 0; j < KEYS.length; j++) {
                    var optKey = KEYS[j];
                    if (!Object.prototype.hasOwnProperty.call(urlValues, optKey)) {
                        safeSetStorage(optKey, '');
                    }
                }
            }

            for (var m = 0; m < KEYS.length; m++) {
                var storeKey = KEYS[m];
                if (Object.prototype.hasOwnProperty.call(urlValues, storeKey)) {
                    safeSetStorage(storeKey, urlValues[storeKey]);
                }
            }
        }

        function resolve(key) {
            if (Object.prototype.hasOwnProperty.call(urlValues, key)) {
                return urlValues[key];
            }
            return safeGetStorage(key);
        }

        var dealer = resolve('dealer');
        var contact = resolve('contact');
        var dp = resolve('dp');
        var date = resolve('date');
        var declinesRaw = resolve('declines');
        var marginRaw = resolve('margin');
        var banksRaw = resolve('banks');
        var concernRaw = resolve('concern').toLowerCase();

        var parsedDeclines = parseInt(declinesRaw, 10);
        var declines = (!isNaN(parsedDeclines) && parsedDeclines > 0) ? parsedDeclines : null;

        var parsedMargin = parseFloat(marginRaw);
        var marginK = null;
        if (!isNaN(parsedMargin) && parsedMargin > 0) {
            marginK = parsedMargin >= 1000 ? Math.round((parsedMargin / 1000) * 10) / 10 : parsedMargin;
        }

        var banks = banksRaw
            ? banksRaw.split(',').map(function (b) { return b.trim(); }).filter(Boolean).join(', ')
            : '';

        var concern = VALID_CONCERNS[concernRaw] ? concernRaw : '';
        var displayDealer = dealer || 'Your Dealership';

        return {
            dealer: dealer,
            displayDealer: displayDealer,
            dealerSlug: toDealerSlug(displayDealer),
            contact: contact,
            dp: dp,
            date: date,
            declines: declines,
            declinesForMath: declines || 25,
            marginK: marginK,
            banks: banks,
            concern: concern,
            hasUrlParams: hasUrlParams
        };
    }

    function buildQueryString(prospect) {
        var qs = new URLSearchParams();
        if (prospect.dealer) qs.set('dealer', prospect.dealer);
        if (prospect.contact) qs.set('contact', prospect.contact);
        if (prospect.dp) qs.set('dp', prospect.dp);
        if (prospect.date) qs.set('date', prospect.date);
        if (prospect.declines) qs.set('declines', String(prospect.declines));
        if (prospect.marginK) qs.set('margin', String(prospect.marginK));
        if (prospect.banks) qs.set('banks', prospect.banks);
        if (prospect.concern) qs.set('concern', prospect.concern);
        var encoded = qs.toString();
        return encoded ? '?' + encoded : '';
    }

    function applyDomPersonalization(prospect) {
        var dealerIds = [
            'header-dealership-name',
            'hero-dealership-name',
            'footer-dealership-name',
            'meta-dealership'
        ];
        for (var i = 0; i < dealerIds.length; i++) {
            var el = document.getElementById(dealerIds[i]);
            if (el) el.textContent = prospect.displayDealer;
        }

        var dealerClassSelectors = '.dealer-display-name, .dealer-name-insert, .analyst-dealer-label, .modal-dealer-name-footer';
        document.querySelectorAll(dealerClassSelectors).forEach(function (node) {
            node.textContent = prospect.displayDealer;
        });

        document.querySelectorAll('.dealer-slug-insert').forEach(function (node) {
            node.textContent = prospect.dealerSlug;
        });

        var dealerInput = document.getElementById('dealership_name_input');
        if (dealerInput) {
            dealerInput.value = prospect.displayDealer;
        }

        // Propagate ALL active params across EVERY internal .html link while preserving #hash anchors
        var queryStr = buildQueryString(prospect);
        if (queryStr) {
            document.querySelectorAll('a[href*=".html"]').forEach(function (a) {
                var href = a.getAttribute('href');
                if (!href || href.startsWith('http') || href.startsWith('//') || href.startsWith('mailto:') || href.startsWith('tel:')) {
                    return;
                }
                var hashParts = href.split('#');
                var basePath = hashParts[0].split('?')[0];
                var hashSuffix = hashParts.length > 1 ? '#' + hashParts.slice(1).join('#') : '';
                a.setAttribute('href', basePath + queryStr + hashSuffix);
            });
        }
    }

    var prospect = parseProspect();
    window.CBF_PROSPECT = prospect;
    window.CBF_PERSONALIZE = {
        apply: function () {
            applyDomPersonalization(window.CBF_PROSPECT);
        },
        buildQueryString: function () {
            return buildQueryString(window.CBF_PROSPECT);
        }
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function () {
            applyDomPersonalization(window.CBF_PROSPECT);
        });
    } else {
        applyDomPersonalization(window.CBF_PROSPECT);
    }
})();
