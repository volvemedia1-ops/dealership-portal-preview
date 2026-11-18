/**
 * Credit Builder Flow • Dealership Portal Prospect Intelligence & Analytics
 * Integrated with Microsoft Clarity (Project ID: yi9wab3vxv)
 */

(function() {
    'use strict';

    var KEYS = ['dealer', 'contact', 'dp', 'date', 'declines', 'margin', 'banks', 'concern'];
    var STALE_DEFAULTS = {
        'Volkswagen Auto Group': true,
        'Johan Van Der Merwe': true,
        'Johan Botha': true,
        'Apex Auto Group': true
    };

    function sanitizeParam(val) {
        if (val === null || val === undefined) return '';
        var str = String(val).replace(/\+/g, ' ').trim();
        return STALE_DEFAULTS[str] ? '' : str;
    }

    // 1. Prospect Data Extraction (read-only; portal-personalize.js manages persistence)
    function getProspectInfo() {
        try {
            if (window.CBF_PROSPECT) {
                return {
                    dealer: window.CBF_PROSPECT.displayDealer || 'Your Dealership',
                    contact: window.CBF_PROSPECT.contact || window.CBF_PROSPECT.dp || 'Executive Team',
                    declines: String(window.CBF_PROSPECT.declinesForMath || '25'),
                    concern: window.CBF_PROSPECT.concern || ''
                };
            }

            var params = new URLSearchParams(window.location.search);
            var resetVal = (params.get('reset') || '').toLowerCase();
            var isReset = (resetVal === '1' || resetVal === 'true');

            function readField(key) {
                if (params.has(key)) return sanitizeParam(params.get(key));
                if (isReset) return '';
                return sanitizeParam(localStorage.getItem('cbf_' + key));
            }

            var dealer = readField('dealer') || 'Your Dealership';
            var contact = readField('contact') || readField('dp') || 'Executive Team';
            var declines = readField('declines') || '25';
            var concern = readField('concern') || '';

            return {
                dealer: dealer,
                contact: contact,
                declines: declines,
                concern: concern
            };
        } catch (e) {
            return { dealer: 'Your Dealership', contact: 'Executive Team', declines: '25', concern: '' };
        }
    }

    // 2. Identify and Tag Session in Clarity
    function tagClaritySession() {
        if (typeof window.clarity !== 'function') return;

        var prospect = getProspectInfo();

        // Set custom tags searchable and filterable in Clarity Dashboard
        window.clarity('set', 'Dealership', prospect.dealer);
        window.clarity('set', 'ExecutiveContact', prospect.contact);
        window.clarity('set', 'MonthlyDeclines', prospect.declines);
        if (prospect.concern) {
            window.clarity('set', 'PrimaryConcern', prospect.concern);
        }
        window.clarity('set', 'PageTitle', document.title);

        // Identify prospect friendly name for session recordings
        window.clarity('identify', prospect.dealer, null, null, prospect.contact + ' (' + prospect.dealer + ')');
    }

    // Run identification when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', tagClaritySession);
    } else {
        tagClaritySession();
    }

    // 3. High-Intent Sales Event Listeners
    document.addEventListener('click', function(e) {
        if (typeof window.clarity !== 'function') return;

        var el = e.target.closest('a, button');
        if (!el) return;

        var href = el.getAttribute('href') || '';
        var text = (el.innerText || '').toLowerCase();

        // WhatsApp Direct Contact Clicks
        if (href.indexOf('wa.me') !== -1) {
            window.clarity('event', 'click_whatsapp_contact');
        }
        // Online MSA Execution Clicks
        else if (href.indexOf('agreement.volvemedia.com') !== -1 || text.indexOf('review & sign') !== -1) {
            window.clarity('event', 'click_sign_agreement');
        }
        // Dealership Portal Registration Clicks
        else if (href.indexOf('creditbuilderflow.co.za/dealerships') !== -1) {
            window.clarity('event', 'click_register_dealership_portal');
        }
        // Next Steps / Activation Path Clicks
        else if (href.indexOf('05-next-steps') !== -1 || href.indexOf('activation-steps') !== -1) {
            window.clarity('event', 'click_next_steps');
        }
        // Master Service Agreement Inspection
        else if (href.indexOf('dealership-service-agreement') !== -1 || text.indexOf('agreement') !== -1) {
            window.clarity('event', 'click_view_agreement');
        }
        // Forensic Case Study Inspections
        else if (href.indexOf('case-study-') !== -1) {
            window.clarity('event', 'click_case_study_dossier');
        }
        // Start Recovery CTAs
        else if (text.indexOf('start recovering') !== -1 || text.indexOf('activate cbf') !== -1) {
            window.clarity('event', 'click_start_recovering');
        }
    });

    // 4. ROI Calculator Interaction Tracking (if on 01-commercial-roi.html)
    document.addEventListener('DOMContentLoaded', function() {
        var leadsSlider = document.getElementById('leads') || document.getElementById('declines-slider');
        var profitSlider = document.getElementById('profit') || document.getElementById('profit-slider');
        var vehiclePriceSlider = document.getElementById('vehicle-price');

        if (leadsSlider) {
            leadsSlider.addEventListener('change', function() {
                if (typeof window.clarity === 'function') {
                    window.clarity('event', 'adjusted_declines_slider');
                    window.clarity('set', 'CalculatedDeclines', leadsSlider.value);
                }
            });
        }

        if (profitSlider) {
            profitSlider.addEventListener('change', function() {
                if (typeof window.clarity === 'function') {
                    window.clarity('event', 'adjusted_profit_slider');
                    window.clarity('set', 'CalculatedProfitPerUnit', profitSlider.value);
                }
            });
        }

        if (vehiclePriceSlider) {
            vehiclePriceSlider.addEventListener('change', function() {
                if (typeof window.clarity === 'function') {
                    window.clarity('event', 'adjusted_vehicle_price_slider');
                }
            });
        }
    });

    // Global helper
    window.cbfTrack = function(eventName, tagData) {
        if (typeof window.clarity === 'function') {
            if (eventName) window.clarity('event', eventName);
            if (tagData && typeof tagData === 'object') {
                for (var key in tagData) {
                    if (tagData.hasOwnProperty(key)) {
                        window.clarity('set', key, String(tagData[key]));
                    }
                }
            }
        }
    };
})();
