/**
 * Zotero QuickCite - Local Zotero Desktop API Client
 * Uses ONLYOFFICE native AscSimpleRequest (C++ backend) to bypass Chromium CORS/Origin restrictions.
 * Decodes all HTML entities and strips Unicode isolate characters so citations match Zotero 100%.
 * Features agentic multi-tier search fallback to guarantee maximum recall across all query combinations.
 */
var ZoteroAPI = (function () {
    "use strict";

    var BASE_URL = "http://127.0.0.1:23119";
    var ZOTERO_API_VERSION = "3";
    var USER_AGENT = "AscDesktopEditor";

    var _isConnected = false;
    var _userGroups = [];

    /**
     * Decode HTML entities and strip Unicode directional isolates/zero-width chars
     */
    function decodeHtmlAndClean(str) {
        if (!str) return "";
        return str
            .replace(/&#38;/g, "&")
            .replace(/&amp;/g, "&")
            .replace(/&#60;/g, "<")
            .replace(/&lt;/g, "<")
            .replace(/&#62;/g, ">")
            .replace(/&gt;/g, ">")
            .replace(/[\u2066-\u206f\u200b-\u200f]/g, "") // strip U+2068, U+2069, etc.
            .replace(/\s+/g, " ")
            .trim();
    }

    /**
     * Locate or initialize AscSimpleRequest across available window contexts
     */
    function getSimpleRequestClient() {
        if (window.AscSimpleRequest && typeof window.AscSimpleRequest.createRequest === "function") {
            return window.AscSimpleRequest;
        }
        if (window.parent && window.parent.AscSimpleRequest && typeof window.parent.AscSimpleRequest.createRequest === "function") {
            return window.parent.AscSimpleRequest;
        }
        if (window.top && window.top.AscSimpleRequest && typeof window.top.AscSimpleRequest.createRequest === "function") {
            return window.top.AscSimpleRequest;
        }

        var desktop = window.AscDesktopEditor ||
                      (window.parent && window.parent.AscDesktopEditor) ||
                      (window.top && window.top.AscDesktopEditor);

        if (desktop && typeof desktop.sendSimpleRequest === "function") {
            var reqCounter = 0;
            var pending = {};

            window.AscSimpleRequest = {
                createRequest: function (cfg) {
                    var id = ++reqCounter;
                    pending[id] = { complete: cfg.complete, error: cfg.error };
                    if (cfg.timeout) {
                        setTimeout(function () {
                            if (pending[id]) {
                                if (pending[id].error) {
                                    pending[id].error({ statusCode: 404, message: "Request timed out" });
                                }
                                delete pending[id];
                            }
                        }, cfg.timeout);
                    }
                    desktop.sendSimpleRequest(id, cfg);
                },
                _onSuccess: function (id, res) {
                    if (pending[id]) {
                        if (pending[id].complete) pending[id].complete(res);
                        delete pending[id];
                    }
                },
                _onError: function (id, err) {
                    if (pending[id]) {
                        if (pending[id].error) pending[id].error(err);
                        delete pending[id];
                    }
                }
            };
            return window.AscSimpleRequest;
        }

        return null;
    }

    /**
     * Unified HTTP GET request dispatcher with safe 12-second timeout
     */
    function sendGet(url, timeoutMs) {
        timeoutMs = timeoutMs || 12000; // Generous 12s timeout so large queries never prematurely fail
        var client = getSimpleRequestClient();

        if (client) {
            return new Promise(function (resolve, reject) {
                client.createRequest({
                    url: url,
                    method: "GET",
                    headers: {
                        "Zotero-API-Version": ZOTERO_API_VERSION,
                        "User-Agent": USER_AGENT
                    },
                    timeout: timeoutMs,
                    complete: function (res) {
                        var status = (res && res.responseStatus !== undefined) ? res.responseStatus :
                                     ((res && res.statusCode !== undefined) ? res.statusCode : 200);

                        if (status === 200 || status === 201) {
                            resolve(res.responseText || "");
                        } else if (status === 403) {
                            reject(new Error("Permission denied. Enable 'Allow other applications' in Zotero Advanced Settings."));
                        } else {
                            reject(new Error("Zotero responded with HTTP " + status));
                        }
                    },
                    error: function (err) {
                        var msg = (err && err.message) ? err.message : "Connection failed";
                        reject(new Error(msg));
                    }
                });
            });
        }

        return fetch(url, {
            method: "GET",
            headers: {
                "Zotero-API-Version": ZOTERO_API_VERSION,
                "User-Agent": USER_AGENT
            }
        }).then(function (res) {
            if (!res.ok) throw new Error("HTTP " + res.status);
            return res.text();
        });
    }

    /**
     * Ping local Zotero desktop app
     */
    function ping() {
        return sendGet(BASE_URL + "/api/", 2500).then(function () {
            _isConnected = true;
            return true;
        }).catch(function (err) {
            console.warn("Zotero ping error:", err);
            _isConnected = false;
            return false;
        });
    }

    /**
     * Validate whether a style is installed and accepted by local Zotero Desktop
     */
    function validateStyle(styleId) {
        if (!styleId) return Promise.resolve(false);
        var url = BASE_URL + "/api/users/0/items?limit=1&format=bib&style=" + encodeURIComponent(styleId.trim());
        return sendGet(url, 2500).then(function () {
            return true;
        }).catch(function () {
            return false;
        });
    }

    /**
     * Raw single-query execution against Zotero API
     */
    function rawSearch(q, libraryId, style) {
        var lib = (libraryId && libraryId !== "0" && libraryId !== 0) ? ("groups/" + libraryId) : "users/0";
        var cslStyle = style || "apa";

        var url = BASE_URL + "/api/" + lib + "/items?itemType=-attachment&include=citation,data&style=" + encodeURIComponent(cslStyle) + "&limit=15";
        if (q && q.trim().length > 0) {
            url += "&q=" + encodeURIComponent(q.trim());
        }

        return sendGet(url, 12000).then(function (rawText) {
            var data = [];
            try {
                data = JSON.parse(rawText);
            } catch (e) {
                console.warn("Failed to parse items JSON:", e);
            }
            return (data || []).map(normalizeItem);
        });
    }

    /**
     * Agentic Multi-Tier Smart Search:
     * Tier 1: Exact query (e.g. "Yan meta 2026")
     * Tier 2: Year extraction fallback (extracts 4-digit year, queries textual terms, and ranks target year first)
     * Tier 3: Query relaxation fallback for 3+ words (drops auxiliary word to recover relevant hits)
     */
    function searchItems(query, libraryId, style) {
        if (!query || !query.trim()) {
            return rawSearch("", libraryId, style);
        }

        var cleanQ = query.trim();

        // Tier 1: Try exact query
        return rawSearch(cleanQ, libraryId, style).then(function (results) {
            if (results && results.length > 0) {
                return results;
            }

            // Tier 2: Check if query contains a 4-digit year (e.g. 2026 or 2025)
            var yearMatch = cleanQ.match(/\b(19\d\d|20\d\d)\b/);
            if (yearMatch) {
                var targetYear = yearMatch[1];
                var queryNoYear = cleanQ.replace(/\b(19\d\d|20\d\d)\b/g, "").replace(/\s+/g, " ").trim();

                if (queryNoYear) {
                    return rawSearch(queryNoYear, libraryId, style).then(function (noYearResults) {
                        if (noYearResults && noYearResults.length > 0) {
                            var matchingYear = noYearResults.filter(function (it) {
                                return it.year === targetYear || (it.raw && it.raw.data && String(it.raw.data.date).indexOf(targetYear) !== -1);
                            });
                            var otherYears = noYearResults.filter(function (it) {
                                return it.year !== targetYear && (!it.raw || !it.raw.data || String(it.raw.data.date).indexOf(targetYear) === -1);
                            });
                            return matchingYear.concat(otherYears);
                        }
                        return tryRelaxedSearch(cleanQ, libraryId, style);
                    });
                }
            }

            // Tier 3: Query relaxation for 3+ words
            return tryRelaxedSearch(cleanQ, libraryId, style);
        });
    }

    function tryRelaxedSearch(cleanQ, libraryId, style) {
        var words = cleanQ.split(/\s+/);
        if (words.length >= 3) {
            var relaxed = words.slice(0, Math.min(3, words.length - 1)).join(" ");
            return rawSearch(relaxed, libraryId, style).then(function (relaxedResults) {
                return relaxedResults || [];
            });
        }
        return Promise.resolve([]);
    }

    /**
     * Get item details by keys
     */
    function getItemsByKeys(keys, libraryId, style) {
        if (!keys || !keys.length) return Promise.resolve([]);
        var lib = (libraryId && libraryId !== "0" && libraryId !== 0) ? ("groups/" + libraryId) : "users/0";
        var cslStyle = style || "apa";

        var url = BASE_URL + "/api/" + lib + "/items?itemKey=" + encodeURIComponent(keys.join(",")) + "&itemType=-attachment&include=citation,data&style=" + encodeURIComponent(cslStyle);

        return sendGet(url, 8000).then(function (rawText) {
            var data = [];
            try {
                data = JSON.parse(rawText);
            } catch (e) {
                console.warn("Failed to parse items JSON:", e);
            }
            return (data || [])
                .filter(function (it) {
                    return it && (!it.data || it.data.itemType !== "attachment");
                })
                .map(normalizeItem);
        });
    }

    /**
     * Get formatted bibliography for keys
     */
    function getBibliography(keys, libraryId, style) {
        if (!keys || !keys.length) return Promise.resolve("");
        var lib = (libraryId && libraryId !== "0" && libraryId !== 0) ? ("groups/" + libraryId) : "users/0";
        var cslStyle = style || "apa";

        var url = BASE_URL + "/api/" + lib + "/items?itemKey=" + encodeURIComponent(keys.join(",")) + "&itemType=-attachment&format=bib&style=" + encodeURIComponent(cslStyle);

        return sendGet(url, 12000);
    }

    /**
     * Fetch user group libraries
     */
    function getGroups() {
        return sendGet(BASE_URL + "/api/users/0/groups", 3000).then(function (rawText) {
            var data = [];
            try {
                data = JSON.parse(rawText);
            } catch (e) {}
            _userGroups = (data || []).map(function (g) {
                return {
                    id: g.id,
                    name: g.data ? g.data.name : ("Group " + g.id)
                };
            });
            return _userGroups;
        }).catch(function () {
            return [];
        });
    }

    /**
     * Normalize raw Zotero API item into structured object
     * Strips Unicode isolate characters and extracts clean authors and Zotero's exact citation
     */
    function normalizeItem(raw) {
        var data = raw.data || {};
        var meta = raw.meta || {};

        var authors = "";
        if (Array.isArray(data.creators) && data.creators.length > 0) {
            var first = data.creators[0];
            var lastName = (first.lastName || first.name || "").trim();
            if (data.creators.length === 1) {
                authors = lastName;
            } else if (data.creators.length === 2) {
                var second = data.creators[1];
                var secondLast = (second.lastName || second.name || "").trim();
                authors = lastName + " & " + secondLast;
            } else {
                authors = lastName + " et al.";
            }
        } else if (meta.creatorSummary) {
            authors = decodeHtmlAndClean(meta.creatorSummary);
        }

        var year = "";
        if (meta.parsedDate) {
            year = meta.parsedDate.substring(0, 4);
        } else if (data.date) {
            var m = data.date.match(/\b(19\d\d|20\d\d)\b/);
            if (m) year = m[1];
            else year = data.date;
        }

        var citationText = decodeHtmlAndClean(raw.citation || "");

        return {
            key: raw.key || data.key,
            id: raw.key || data.key,
            itemType: data.itemType || "journalArticle",
            title: decodeHtmlAndClean(data.title || "[No title]"),
            authors: authors,
            year: year,
            publication: decodeHtmlAndClean(data.publicationTitle || data.journalAbbreviation || data.bookTitle || data.publisher || ""),
            citation: citationText,
            raw: raw
        };
    }

    return {
        ping: ping,
        validateStyle: validateStyle,
        searchItems: searchItems,
        getItemsByKeys: getItemsByKeys,
        getBibliography: getBibliography,
        getGroups: getGroups,
        decodeHtmlAndClean: decodeHtmlAndClean,
        isConnected: function () { return _isConnected; }
    };
})();
