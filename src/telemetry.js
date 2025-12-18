/*
 * Telemetry module for Outlook Recipient Warning Add-in
 * Uses Azure Application Insights for usage analytics and error tracking
 * All methods are fail-safe - telemetry errors never affect add-in functionality
 * ES2016 compatible (no async/await, no optional chaining)
 */

var Telemetry = (function() {
    "use strict";

    var appInsights = null;
    var isInitialized = false;
    var CONNECTION_STRING = "InstrumentationKey=8fd46d6e-1a20-464a-9038-ac0998e1b79c;IngestionEndpoint=https://norwayeast-0.in.applicationinsights.azure.com/;LiveEndpoint=https://norwayeast.livediagnostics.monitor.azure.com/;ApplicationId=d37df70d-8612-4476-adab-c45cfd64d749";

    /**
     * Initialize Application Insights
     * Call this once when the add-in loads
     */
    function initialize() {
        try {
            if (isInitialized) {
                return;
            }

            if (typeof Microsoft !== "undefined" &&
                Microsoft.ApplicationInsights &&
                Microsoft.ApplicationInsights.ApplicationInsights) {

                appInsights = new Microsoft.ApplicationInsights.ApplicationInsights({
                    config: {
                        connectionString: CONNECTION_STRING,
                        enableAutoRouteTracking: false,
                        disableFetchTracking: false,
                        disableAjaxTracking: false,
                        enableCorsCorrelation: true,
                        enableRequestHeaderTracking: true,
                        enableResponseHeaderTracking: true
                    }
                });
                appInsights.loadAppInsights();

                // Track unhandled errors automatically
                appInsights.addTelemetryInitializer(function(envelope) {
                    // Add custom properties to all telemetry
                    if (envelope.data) {
                        envelope.data.addinVersion = "2.5.0.0";
                        envelope.data.addinName = "RecipientPrivacyWarning";
                    }
                    return true;
                });

                isInitialized = true;
            }
        } catch (e) {
            // Silent fail - telemetry should never break the add-in
        }
    }

    /**
     * Track a custom event
     * @param {string} name - Event name
     * @param {Object} properties - Optional custom properties
     * @param {Object} measurements - Optional numeric measurements
     */
    function trackEvent(name, properties, measurements) {
        try {
            if (!isInitialized || !appInsights) {
                return;
            }
            appInsights.trackEvent({
                name: name,
                properties: properties,
                measurements: measurements
            });
        } catch (e) {
            // Silent fail
        }
    }

    /**
     * Track an exception/error
     * @param {Error} error - The error object
     * @param {Object} properties - Optional custom properties
     */
    function trackException(error, properties) {
        try {
            if (!isInitialized || !appInsights) {
                return;
            }
            appInsights.trackException({
                exception: error,
                properties: properties
            });
        } catch (e) {
            // Silent fail
        }
    }

    /**
     * Track a page view (task pane open)
     * @param {string} name - Page name
     * @param {Object} properties - Optional custom properties
     */
    function trackPageView(name, properties) {
        try {
            if (!isInitialized || !appInsights) {
                return;
            }
            appInsights.trackPageView({
                name: name,
                properties: properties
            });
        } catch (e) {
            // Silent fail
        }
    }

    /**
     * Track a metric/measurement
     * @param {string} name - Metric name
     * @param {number} value - Metric value
     * @param {Object} properties - Optional custom properties
     */
    function trackMetric(name, value, properties) {
        try {
            if (!isInitialized || !appInsights) {
                return;
            }
            appInsights.trackMetric({
                name: name,
                average: value,
                properties: properties
            });
        } catch (e) {
            // Silent fail
        }
    }

    /**
     * Flush telemetry data (call before page unload if needed)
     */
    function flush() {
        try {
            if (!isInitialized || !appInsights) {
                return;
            }
            appInsights.flush();
        } catch (e) {
            // Silent fail
        }
    }

    // Public API
    return {
        initialize: initialize,
        trackEvent: trackEvent,
        trackException: trackException,
        trackPageView: trackPageView,
        trackMetric: trackMetric,
        flush: flush
    };
})();
