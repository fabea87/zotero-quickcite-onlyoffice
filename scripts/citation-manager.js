/**
 * Zotero QuickCite - Citation & Document Manager
 * Full Word-compatible AddinField and Document Property integration.
 * Compatible with files created or edited by Microsoft Word official Zotero plugin,
 * CSL formatting, hanging indents (悬挂缩进), ubiquitous narrative citations,
 * automatic bibliography upgrades, and document-wide style changes in Times New Roman.
 */
var CitationManager = (function () {
    "use strict";

    var CIT_PREFIX = "ZOTERO_ITEM CSL_CITATION";
    var BIB_PREFIX = "ZOTERO_BIBL {\"uncited\":[],\"omitted\":[],\"custom\":[]} CSL_BIBLIOGRAPHY";

    function generateId() {
        return Math.random().toString(36).substring(2, 10);
    }

    /**
     * Decode HTML entities and clean Unicode directional isolates/zero-width chars
     */
    function cleanText(str) {
        if (!str) return "";
        return str
            .replace(/&#38;/g, "&")
            .replace(/&amp;/g, "&")
            .replace(/&#60;/g, "<")
            .replace(/&lt;/g, "<")
            .replace(/&#62;/g, ">")
            .replace(/&gt;/g, ">")
            .replace(/[\u2066-\u206f\u200b-\u200f]/g, "")
            .replace(/\s+/g, " ")
            .trim();
    }

    /**
     * Resolve the alphanumeric Zotero item key from any citation item structure:
     * - item.key
     * - item.uris array (e.g. "http://zotero.org/users/8437599/items/VTPUUB69")
     * - item.uri
     * - 8-char alphanumeric item.id
     * - item.itemData.key
     */
    function getItemKey(item) {
        if (!item) return "";
        if (item.key) return String(item.key);
        if (Array.isArray(item.uris) && item.uris.length > 0) {
            for (var i = 0; i < item.uris.length; i++) {
                var m = String(item.uris[i]).match(/\/items\/([A-Za-z0-9]+)/);
                if (m && m[1]) return m[1];
            }
        }
        if (item.uri) {
            var m2 = String(item.uri).match(/\/items\/([A-Za-z0-9]+)/);
            if (m2 && m2[1]) return m2[1];
        }
        if (typeof item.id === "string" && /^[A-Za-z0-9]{8}$/.test(item.id)) {
            return item.id;
        }
        if (item.itemData && item.itemData.key) {
            return String(item.itemData.key);
        }
        return item.id ? String(item.id) : "";
    }

    /**
     * Fallback item reconstructor from CSL-JSON itemData (for items cited from external libraries or offline)
     */
    function createItemFromItemData(citItem, key) {
        var d = (citItem && citItem.itemData) || {};
        var authors = "";
        if (Array.isArray(d.author) && d.author.length > 0) {
            var first = d.author[0];
            var lastName = (first.family || first.name || "").trim();
            if (d.author.length === 1) {
                authors = lastName;
            } else if (d.author.length === 2) {
                var second = d.author[1];
                authors = lastName + " & " + (second.family || second.name || "").trim();
            } else {
                authors = lastName + " et al.";
            }
        } else if (Array.isArray(d.creators) && d.creators.length > 0) {
            var firstC = d.creators[0];
            var lastC = (firstC.lastName || firstC.name || "").trim();
            if (d.creators.length === 1) {
                authors = lastC;
            } else if (d.creators.length === 2) {
                var secondC = d.creators[1];
                authors = lastC + " & " + (secondC.lastName || secondC.name || "").trim();
            } else {
                authors = lastC + " et al.";
            }
        }

        var year = "";
        if (d.issued && d.issued["date-parts"] && d.issued["date-parts"][0]) {
            year = String(d.issued["date-parts"][0][0] || "");
        } else if (d.date) {
            var my = String(d.date).match(/\b(19\d\d|20\d\d)\b/);
            year = my ? my[1] : String(d.date);
        }

        var k = key || getItemKey(citItem) || String(citItem.id || "");
        return {
            key: k,
            id: k,
            itemType: d.type || d.itemType || "journalArticle",
            title: cleanText(d.title || "[No title]"),
            authors: authors || "Unknown",
            year: year,
            publication: cleanText(d["container-title"] || d.publicationTitle || d.publisher || ""),
            citation: (citItem && (citItem.formattedCitation || (citItem.properties && citItem.properties.formattedCitation))) || "",
            raw: { data: d },
            itemData: d,
            uris: citItem.uris || []
        };
    }

    /**
     * Format locator according to academic CSL rules:
     * - Single page: "p. 42"
     * - Page range (contains -, –, ,, or 'to'): "pp. 42-44"
     * - If already prefixed with p., pp., chap., sec., etc., leave intact
     */
    function formatLocator(rawLoc) {
        if (!rawLoc) return "";
        var loc = rawLoc.trim();
        if (!loc) return "";
        if (/^(p\.|pp\.|chap\.|sec\.|para\.)/i.test(loc)) return loc;
        if (/[\-\u2013,]|(\bto\b)/i.test(loc)) return "pp. " + loc;
        return "p. " + loc;
    }

    /**
     * Get the addin field at current cursor position (with Bookmark mode fallback)
     */
    function getCurrentField() {
        return new Promise(function (resolve) {
            try {
                window.Asc.plugin.executeMethod("GetCurrentAddinField", null, function (field) {
                    if (field && field.Value && field.Value.indexOf("ZOTERO_ITEM") !== -1) {
                        resolve(parseCitationField(field));
                        return;
                    }
                    // Bookmark mode fallback (e.g. ZOTERO_BREF_ from Word/LibreOffice/Google Docs)
                    window.Asc.plugin.executeMethod("GetCurrentBookmark", null, function (bmarkName) {
                        if (bmarkName && bmarkName.indexOf("ZOTERO_BREF_") === 0) {
                            window.Asc.scope.targetBmark = bmarkName;
                            window.Asc.plugin.callCommand(function () {
                                var bm = Asc.scope.targetBmark;
                                var doc = Api.GetDocument();
                                var props = doc.GetCustomProperties ? doc.GetCustomProperties() : null;
                                var val = props && props.Get ? props.Get(bm) : "";
                                var rng = doc.GetBookmarkRange ? doc.GetBookmarkRange(bm) : null;
                                var txt = rng && rng.GetText ? rng.GetText() : "";
                                return { FieldId: bm, Value: val, Content: txt, isBookmark: true };
                            }, false, true, function (res) {
                                if (res && res.Value) {
                                    resolve(parseCitationField(res));
                                } else {
                                    resolve(null);
                                }
                            });
                        } else {
                            resolve(null);
                        }
                    });
                });
            } catch (e) {
                console.warn("GetCurrentAddinField failed:", e);
                resolve(null);
            }
        });
    }

    /**
     * Get all Zotero citation fields in document (both AddinFields and ZOTERO_BREF_ Bookmarks)
     * Queries AddinFields and Bookmarks concurrently for maximum responsiveness.
     */
    function getAllCitationFields() {
        var pAddin = new Promise(function (res) {
            try {
                window.Asc.plugin.executeMethod("GetAllAddinFields", null, function (fields) {
                    var zoteroFields = [];
                    if (Array.isArray(fields)) {
                        fields.forEach(function (f) {
                            if (f && f.Value && f.Value.indexOf("ZOTERO_ITEM") !== -1) {
                                var parsed = parseCitationField(f);
                                if (parsed) zoteroFields.push(parsed);
                            }
                        });
                    }
                    res(zoteroFields);
                });
            } catch (e) {
                res([]);
            }
        });

        var pBmarks = new Promise(function (res) {
            try {
                window.Asc.plugin.callCommand(function () {
                    var doc = Api.GetDocument();
                    var bmarks = doc.GetAllBookmarksNames ? doc.GetAllBookmarksNames() : [];
                    var list = [];
                    var props = doc.GetCustomProperties ? doc.GetCustomProperties() : null;
                    if (props && props.Get) {
                        for (var i = 0; i < bmarks.length; i++) {
                            var bm = bmarks[i];
                            if (bm.indexOf("ZOTERO_BREF_") === 0) {
                                var val = props.Get(bm);
                                if (val) {
                                    var rng = doc.GetBookmarkRange ? doc.GetBookmarkRange(bm) : null;
                                    var txt = rng && rng.GetText ? rng.GetText() : "";
                                    list.push({ FieldId: bm, Value: val, Content: txt, isBookmark: true });
                                }
                            }
                        }
                    }
                    return list;
                }, false, true, function (bmarkList) {
                    var parsedList = [];
                    if (Array.isArray(bmarkList)) {
                        bmarkList.forEach(function (bf) {
                            var parsed = parseCitationField(bf);
                            if (parsed) parsedList.push(parsed);
                        });
                    }
                    res(parsedList);
                });
            } catch (e) {
                res([]);
            }
        });

        return Promise.all([pAddin, pBmarks]).then(function (results) {
            return results[0].concat(results[1]);
        });
    }

    /**
     * Locate existing bibliography representation in document:
     * Checks ONLYOFFICE ContentControl (ZOTERO_BIBLIOGRAPHY), Word AddinField (ADDIN ZOTERO_BIBL),
     * or document References/Bibliography heading section.
     */
    function getBibliographyInfo() {
        return new Promise(function (resolve) {
            try {
                window.Asc.plugin.executeMethod("GetAllContentControls", null, function (controls) {
                    if (Array.isArray(controls)) {
                        for (var j = 0; j < controls.length; j++) {
                            var c = controls[j];
                            if (c && (c.Tag === "ZOTERO_BIBLIOGRAPHY" || c.InternalId === "ZOTERO_BIBLIOGRAPHY")) {
                                resolve({ type: "contentControl", control: c });
                                return;
                            }
                        }
                    }

                    window.Asc.plugin.executeMethod("GetAllAddinFields", null, function (fields) {
                        if (Array.isArray(fields)) {
                            for (var i = 0; i < fields.length; i++) {
                                var f = fields[i];
                                if (f && f.Value && (f.Value.indexOf("ZOTERO_BIBL") !== -1 || f.Value.indexOf("ZOTERO_BIBLIOGRAPHY") !== -1)) {
                                    resolve({ type: "addinField", field: f });
                                    return;
                                }
                            }
                        }

                        // Also check if document has a References / Bibliography heading
                        window.Asc.plugin.callCommand(function () {
                            var doc = Api.GetDocument();
                            var paras = doc.GetAllParagraphs();
                            for (var p = 0; p < paras.length; p++) {
                                var txt = paras[p].GetText().trim();
                                if (/^(references|bibliography|参考文献)\b/i.test(txt)) {
                                    return true;
                                }
                            }
                            return false;
                        }, false, true, function (hasRefSection) {
                            if (hasRefSection) {
                                resolve({ type: "headingSection" });
                            } else {
                                resolve(null);
                            }
                        });
                    });
                });
            } catch (e) {
                console.warn("getBibliographyInfo error:", e);
                resolve(null);
            }
        });
    }

    /**
     * Check if a bibliography exists in the document
     */
    function hasBibliography() {
        return getBibliographyInfo().then(function (info) {
            return info !== null;
        });
    }

    /**
     * Read embedded Zotero style preference from Word custom document properties (ZOTERO_PREF_1 / ZOTERO_PREF_2)
     */
    function getDocumentZoteroStyle() {
        return new Promise(function (resolve) {
            try {
                window.Asc.plugin.callCommand(function () {
                    var result = null;
                    try {
                        var doc = Api.GetDocument();
                        if (doc && doc.GetCustomProperties) {
                            var props = doc.GetCustomProperties();
                            if (props && props.Get) {
                                var p1 = props.Get("ZOTERO_PREF_1") || "";
                                var p2 = props.Get("ZOTERO_PREF_2") || "";
                                var combined = String(p1) + String(p2);
                                var m = combined.match(/style\s+id=["'](?:http:\/\/www\.zotero\.org\/styles\/)?([^"']+)["']/i);
                                if (m && m[1]) {
                                    result = m[1];
                                }
                            }
                        }
                    } catch (e) {}
                    return result;
                }, false, true, function (styleId) {
                    resolve(styleId || null);
                });
            } catch (e) {
                resolve(null);
            }
        });
    }

    /**
     * Save embedded Zotero style preference to Word custom document properties (ZOTERO_PREF_1 / ZOTERO_PREF_2)
     * Ensures complete bidirectional compatibility when reopened in Word with official Zotero plugin.
     */
    function setDocumentZoteroStyle(styleId) {
        if (!styleId) return Promise.resolve();
        return new Promise(function (resolve) {
            try {
                window.Asc.scope.targetStyleId = styleId;
                window.Asc.plugin.callCommand(function () {
                    try {
                        var doc = Api.GetDocument();
                        if (doc && doc.GetCustomProperties) {
                            var props = doc.GetCustomProperties();
                            if (props && props.Add) {
                                var cleanId = Asc.scope.targetStyleId.replace(/^http:\/\/www\.zotero\.org\/styles\//, "").trim();
                                var fullStyleUrl = "http://www.zotero.org/styles/" + cleanId;
                                var sessionId = Math.random().toString(36).substring(2, 10);
                                var pref = '<data data-version="3" zotero-version="10.0.7"><session id="' + sessionId + '"/><style id="' + fullStyleUrl + '" locale="en-US" hasBibliography="1" bibliographyStyleHasBeenSet="1"/><prefs><pref name="fieldType" value="Field"/><pref name="automaticJournalAbbreviations" value="true"/></prefs></data>';
                                if (pref.length > 250) {
                                    props.Add("ZOTERO_PREF_1", pref.substring(0, 250));
                                    props.Add("ZOTERO_PREF_2", pref.substring(250));
                                } else {
                                    props.Add("ZOTERO_PREF_1", pref);
                                }
                            }
                        }
                    } catch (e) {}
                }, false, true, function () {
                    resolve();
                });
            } catch (e) {
                resolve();
            }
        });
    }

    /**
     * Parse raw field into structured citation object.
     * Safely slices JSON between first '{' and last '}' to handle both ONLYOFFICE and MS Word 'ADDIN ...' field prefixes.
     */
    function parseCitationField(rawField) {
        if (!rawField || !rawField.Value) return null;

        var val = rawField.Value;
        var startIdx = val.indexOf("{");
        var endIdx = val.lastIndexOf("}");
        var data = {};
        if (startIdx !== -1 && endIdx !== -1 && endIdx > startIdx) {
            var jsonStr = val.substring(startIdx, endIdx + 1);
            try {
                data = JSON.parse(jsonStr);
            } catch (e) {
                console.warn("Could not parse citation JSON:", jsonStr, e);
            }
        }

        return {
            fieldId: rawField.FieldId,
            rawField: rawField,
            citationID: data.citationID || generateId(),
            properties: data.properties || {},
            citationItems: data.citationItems || [],
            content: rawField.Content || ""
        };
    }

    /**
     * Pure style-based narrative citation generator:
     * Derives author and year representations directly and exclusively from Zotero's CSL output.
     * Zero hardcoded language heuristics; 100% regulated by the style hosted in Zotero.
     */
    function extractNarrativeFromZoteroStyle(zoteroCit, year, locator, prefix, suffix) {
        var raw = cleanText(zoteroCit || "");
        var inner = raw;
        if ((inner.startsWith("(") && inner.endsWith(")")) || (inner.startsWith("[") && inner.endsWith("]"))) {
            inner = inner.slice(1, -1).trim();
        }

        var authorPart = "";
        var yearPart = year || "";

        // Match "<AuthorPart>, <YearPart>" or "<AuthorPart> <YearPart>" or "<AuthorPart>: <YearPart>"
        var m = inner.match(/^(.*?)(?:,\s*|\s+|:\s*)(\d{4}[a-z]?)$/);
        if (m) {
            authorPart = m[1].trim();
            yearPart = m[2];
        } else {
            authorPart = inner;
        }

        // In running narrative prose, if the style generated an ampersand '&' (e.g. APA 'Teng & Zhang'),
        // spell out 'and' (APA 7th §8.11: in narrative citations, spell out 'and' in running text)
        authorPart = authorPart.replace(/\s*&\s*/g, " and ");

        var inParen = yearPart;
        if (locator) {
            inParen += (inParen ? ", " : "") + formatLocator(locator);
        }
        if (suffix) {
            inParen += (suffix.startsWith(",") ? "" : " ") + suffix.trim();
        }

        var full = authorPart + (inParen ? (" (" + inParen + ")") : "");
        if (prefix) {
            full = prefix.trim() + " " + full;
        }

        return {
            author: authorPart,
            year: yearPart,
            fullText: full
        };
    }

    /**
     * Academic Cluster Auto-Sorting Comparator (APA 7th §8.12, Chicago §15.30):
     * Sorts multiple citations in a cluster alphabetically by first author's surname,
     * then chronologically by publication year, then by title.
     */
    function getPrimaryAuthorLastName(item) {
        if (!item) return "";
        if (item.authors) {
            var first = item.authors.split(/[,&]|\bet al/i)[0].trim().toLowerCase();
            return first;
        }
        if (item.raw && item.raw.data && Array.isArray(item.raw.data.creators) && item.raw.data.creators.length > 0) {
            var c = item.raw.data.creators[0];
            return (c.lastName || c.name || "").trim().toLowerCase();
        }
        if (item.itemData && Array.isArray(item.itemData.author) && item.itemData.author.length > 0) {
            var a = item.itemData.author[0];
            return (a.family || a.name || "").trim().toLowerCase();
        }
        return (item.title || "").trim().toLowerCase();
    }

    function getItemYear(item) {
        if (!item) return 0;
        if (item.year) {
            var m = String(item.year).match(/\b(19\d\d|20\d\d)\b/);
            if (m) return parseInt(m[1], 10);
        }
        if (item.raw && item.raw.meta && item.raw.meta.parsedDate) {
            return parseInt(item.raw.meta.parsedDate.substring(0, 4), 10) || 0;
        }
        return 0;
    }

    function compareCitationClusterItems(a, b) {
        var authorA = getPrimaryAuthorLastName(a);
        var authorB = getPrimaryAuthorLastName(b);
        var cmp = authorA.localeCompare(authorB);
        if (cmp !== 0) return cmp;

        var yearA = getItemYear(a);
        var yearB = getItemYear(b);
        if (yearA !== yearB) return yearA - yearB;

        var titleA = (a.title || "").trim().toLowerCase();
        var titleB = (b.title || "").trim().toLowerCase();
        return titleA.localeCompare(titleB);
    }

    /**
     * Build formatted citation text and Word CSL data structure.
     * Auto-sorts multiple citations in parenthetical/suppress-author clusters (APA 7th §8.12).
     * Preserves CSL-JSON itemData, URIs, locators, prefixes, and suffixes.
     */
    function buildCitation(items, options) {
        options = options || {};
        var locator = (options.locator || "").trim();
        var prefix = options.prefix || "";
        var suffix = options.suffix || "";
        var formatMode = options.formatMode || "parenthetical";

        var isNarrative = (formatMode === "narrative");
        var isSuppress = (formatMode === "suppress");

        var narrativeAuthor = "";
        var formattedText = "";

        // Standard Academic Cluster Auto-Sorting:
        // In parenthetical or suppress-author mode, auto-sort cluster items
        // alphabetically by author and chronologically by year (APA 7th §8.12, Chicago §15.30)
        var clusterItems = items.map(function (it, idx) {
            var clone = Object.assign({}, it);
            if (idx === 0) {
                if (locator) clone.locator = locator;
                if (prefix) clone.prefix = prefix;
                if (suffix) clone.suffix = suffix;
            }
            return clone;
        });

        if (!isNarrative && clusterItems.length > 1 && !options.manualOrder) {
            clusterItems.sort(compareCitationClusterItems);
        }

        var citationItems = clusterItems.map(function (item, idx) {
            var isPrimary = (idx === 0);
            var itemLocator = item.locator || "";
            var itemPrefix = item.prefix || "";
            var itemSuffix = item.suffix || "";
            var itemSuppress = isPrimary ? isSuppress : (item["suppress-author"] || false);
            var itemKey = getItemKey(item);

            var itemData = {};
            if (item.itemData && typeof item.itemData === "object" && Object.keys(item.itemData).length > 0) {
                itemData = item.itemData;
            } else if (item.raw && item.raw.data) {
                itemData = item.raw.data;
            }

            var uris = (Array.isArray(item.uris) && item.uris.length > 0)
                ? item.uris
                : ["http://zotero.org/users/0/items/" + itemKey];

            return {
                id: itemKey,
                uris: uris,
                itemData: itemData,
                locator: itemLocator,
                label: itemLocator ? "page" : undefined,
                prefix: itemPrefix,
                suffix: itemSuffix,
                "suppress-author": itemSuppress,
                isNarrative: isNarrative
            };
        });

        if (isNarrative && items.length > 0) {
            var primaryItem = items[0];
            var narrative = extractNarrativeFromZoteroStyle(
                primaryItem.citation,
                primaryItem.year,
                locator,
                prefix,
                suffix
            );
            formattedText = narrative.fullText;
            narrativeAuthor = narrative.author;
        } else if (isSuppress) {
            // YEAR ONLY: (2026) or (2026, pp. 42-44)
            var yearParts = clusterItems.map(function (item, idx) {
                var citItem = citationItems[idx];
                var y = item.year || "";
                if (citItem.locator) {
                    y += ", " + formatLocator(citItem.locator);
                }
                if (citItem.prefix) {
                    y = citItem.prefix.trim() + " " + y;
                }
                if (citItem.suffix) {
                    y = y + (citItem.suffix.startsWith(",") ? "" : " ") + citItem.suffix.trim();
                }
                return y;
            });
            formattedText = "(" + yearParts.join("; ") + ")";
        } else {
            // PARENTHETICAL: DIRECTLY USE ZOTERO DESKTOP'S EXACT CSL OUTPUT
            var isNumeric = false;
            var parts = clusterItems.map(function (item, idx) {
                var citItem = citationItems[idx];
                var zotCit = cleanText(item.citation || "");

                if (!zotCit) {
                    zotCit = "(" + (item.authors || "Unknown") + (item.year ? (", " + item.year) : "") + ")";
                }

                if (zotCit.startsWith("[") && zotCit.endsWith("]")) {
                    isNumeric = true;
                }

                var inner = zotCit;
                if ((inner.startsWith("(") && inner.endsWith(")")) || (inner.startsWith("[") && inner.endsWith("]"))) {
                    inner = inner.slice(1, -1).trim();
                }

                if (citItem.prefix) {
                    inner = citItem.prefix.trim() + " " + inner;
                }
                if (citItem.locator) {
                    var loc = isNumeric ? citItem.locator.trim() : formatLocator(citItem.locator);
                    inner = inner + ", " + loc;
                }
                if (citItem.suffix) {
                    inner = inner + (citItem.suffix.startsWith(",") ? "" : " ") + citItem.suffix.trim();
                }

                return inner;
            });

            if (isNumeric) {
                formattedText = "[" + parts.join(", ") + "]";
            } else {
                formattedText = "(" + parts.join("; ") + ")";
            }
        }

        var citationObj = {
            citationID: options.citationID || generateId(),
            properties: {
                formattedCitation: formattedText,
                plainCitation: formattedText,
                isNarrative: isNarrative,
                noteIndex: options.isFootnote ? 1 : 0
            },
            citationItems: citationItems,
            schema: "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
        };

        return {
            formattedText: formattedText,
            citationObj: citationObj,
            narrativeAuthor: narrativeAuthor,
            isNarrative: isNarrative
        };
    }

    /**
     * Explicitly apply Times New Roman font to a selected field
     */
    function applyFontToField(fieldId) {
        return new Promise(function (resolve) {
            var finished = false;
            var timer = setTimeout(function () {
                if (!finished) {
                    finished = true;
                    resolve();
                }
            }, 600);

            try {
                window.Asc.plugin.executeMethod("SelectAddinField", [fieldId], function () {
                    window.Asc.plugin.callCommand(function () {
                        var oDocument = Api.GetDocument();
                        var oRange = oDocument.GetRangeBySelect();
                        if (oRange && oRange.SetFontFamily) {
                            oRange.SetFontFamily("Times New Roman");
                        }
                    }, false, true, function () {
                        window.Asc.plugin.executeMethod("MoveCursorOutsideField", [fieldId, false], function () {
                            if (!finished) {
                                finished = true;
                                clearTimeout(timer);
                                resolve();
                            }
                        });
                    });
                });
            } catch (e) {
                if (!finished) {
                    finished = true;
                    clearTimeout(timer);
                    resolve();
                }
            }
        });
    }

    /**
     * Pre-fetch updated bibliography HTML from Zotero concurrently with field editing
     */
    function prefetchBibliography(style, extraKeys, bypassCache) {
        return getBibliographyInfo().then(function (bibInfo) {
            if (!bibInfo) return null;

            return getAllCitationFields().then(function (fields) {
                var keysSet = {};
                fields.forEach(function (f) {
                    (f.citationItems || []).forEach(function (it) {
                        var k = getItemKey(it);
                        if (k) keysSet[k] = true;
                    });
                });

                if (Array.isArray(extraKeys)) {
                    extraKeys.forEach(function (k) {
                        if (k) keysSet[k] = true;
                    });
                }

                var keys = Object.keys(keysSet);
                if (!keys.length) return null;

                return ZoteroAPI.getBibliography(keys, 0, style, bypassCache).then(function (htmlBib) {
                    if (!htmlBib) return null;
                    var entries = parseBibHtml(htmlBib);
                    return { htmlBib: htmlBib, entries: entries };
                });
            });
        }).catch(function (e) {
            console.warn("prefetchBibliography error:", e);
            return null;
        });
    }

    /**
     * Insert citation into document with concurrent execution pipeline:
     * Dispatches ONLYOFFICE field insertion and Zotero bibliography retrieval concurrently via Promise.all!
     */
    function insertCitation(citationData, isFootnote, activeStyle) {
        var fieldId = "ZOTERO_" + Date.now();
        var addinField = {
            FieldId: fieldId,
            Value: CIT_PREFIX + " " + JSON.stringify(citationData.citationObj),
            Content: citationData.formattedText
        };

        var newKeys = (citationData.citationObj.citationItems || []).map(function (it) {
            return getItemKey(it);
        });

        function doAddField() {
            if (isFootnote) {
                return new Promise(function (resolve) {
                    window.Asc.plugin.callCommand(function () {
                        var oDocument = Api.GetDocument();
                        oDocument.AddFootnote();
                        var oRange = oDocument.GetRangeBySelect();
                        if (oRange && oRange.SetFontFamily) {
                            oRange.SetFontFamily("Times New Roman");
                        }
                    }, false, true, function () {
                        window.Asc.plugin.executeMethod("AddAddinField", [addinField], function (addedField) {
                            var targetId = (addedField && addedField.FieldId) ? addedField.FieldId : addinField.FieldId;
                            applyFontToField(targetId).then(function () {
                                resolve(addinField);
                            }).catch(function () {
                                resolve(addinField);
                            });
                        });
                    });
                });
            } else {
                return new Promise(function (resolve) {
                    window.Asc.plugin.callCommand(function () {
                        var oDocument = Api.GetDocument();
                        var oRun = oDocument.GetCurrentRun();
                        if (oRun && oRun.SetFontFamily) {
                            oRun.SetFontFamily("Times New Roman");
                        }
                    }, false, true, function () {
                        window.Asc.plugin.executeMethod("AddAddinField", [addinField], function (addedField) {
                            var targetId = (addedField && addedField.FieldId) ? addedField.FieldId : addinField.FieldId;
                            applyFontToField(targetId).then(function () {
                                resolve(addinField);
                            }).catch(function () {
                                resolve(addinField);
                            });
                        });
                    });
                });
            }
        }

        // CONCURRENT PIPELINE: Run ONLYOFFICE insertion and Zotero bibliography prefetch in parallel
        var pInsert = doAddField();
        var pBibPrep = prefetchBibliography(activeStyle, newKeys);

        return Promise.all([pInsert, pBibPrep]).then(function (results) {
            var fieldResult = results[0];
            var bibData = results[1];
            setDocumentZoteroStyle(activeStyle);

            if (bibData && bibData.htmlBib && bibData.entries && bibData.entries.length) {
                return updateBibliographyInDoc(bibData.htmlBib, bibData.entries).then(function () {
                    return fieldResult;
                }).catch(function () {
                    return fieldResult;
                });
            }
            return fieldResult;
        });
    }

    /**
     * Update existing citation in document with concurrent execution pipeline
     * Supports both AddinFields and ZOTERO_BREF_ Bookmarks
     */
    function updateCitation(fieldId, citationData, activeStyle) {
        var isBookmark = (fieldId && String(fieldId).indexOf("ZOTERO_BREF_") === 0);
        var extraKeys = (citationData.citationObj.citationItems || []).map(function (it) {
            return getItemKey(it);
        });

        // CONCURRENT PIPELINE: Prefetch bibliography concurrently with citation update
        var pBibPrep = prefetchBibliography(activeStyle, extraKeys);

        if (isBookmark) {
            var pBmarkUpdate = new Promise(function (resolve) {
                window.Asc.scope.bmarkUpdate = {
                    bmarkName: fieldId,
                    value: CIT_PREFIX + " " + JSON.stringify(citationData.citationObj),
                    text: citationData.formattedText
                };
                window.Asc.plugin.callCommand(function () {
                    var data = Asc.scope.bmarkUpdate;
                    var doc = Api.GetDocument();
                    var props = doc.GetCustomProperties ? doc.GetCustomProperties() : null;
                    if (props && props.Add) {
                        props.Add(data.bmarkName, data.value);
                    }
                    var rng = doc.GetBookmarkRange ? doc.GetBookmarkRange(data.bmarkName) : null;
                    if (rng) {
                        if (rng.SetText) rng.SetText(data.text);
                        if (rng.SetFontFamily) rng.SetFontFamily("Times New Roman");
                    }
                }, false, true, function () {
                    resolve({ FieldId: fieldId, Value: Asc.scope.bmarkUpdate.value, Content: Asc.scope.bmarkUpdate.text });
                });
            });

            return Promise.all([pBmarkUpdate, pBibPrep]).then(function (results) {
                var fieldResult = results[0];
                var bibData = results[1];
                setDocumentZoteroStyle(activeStyle);

                if (bibData && bibData.htmlBib && bibData.entries && bibData.entries.length) {
                    return updateBibliographyInDoc(bibData.htmlBib, bibData.entries).then(function () {
                        return fieldResult;
                    }).catch(function () {
                        return fieldResult;
                    });
                }
                return fieldResult;
            });
        }

        var updatedField = {
            FieldId: fieldId,
            Value: CIT_PREFIX + " " + JSON.stringify(citationData.citationObj),
            Content: citationData.formattedText
        };

        var pUpdate = new Promise(function (resolve) {
            window.Asc.plugin.executeMethod("UpdateAddinFields", [[updatedField]], function () {
                applyFontToField(fieldId).then(function () {
                    resolve(updatedField);
                }).catch(function () {
                    resolve(updatedField);
                });
            });
        });

        return Promise.all([pUpdate, pBibPrep]).then(function (results) {
            var fieldResult = results[0];
            var bibData = results[1];
            setDocumentZoteroStyle(activeStyle);

            if (bibData && bibData.htmlBib && bibData.entries && bibData.entries.length) {
                return updateBibliographyInDoc(bibData.htmlBib, bibData.entries).then(function () {
                    return fieldResult;
                }).catch(function () {
                    return fieldResult;
                });
            }
            return fieldResult;
        });
    }

    /**
     * Remove citation field from document, then awaits auto-upgrading bibliography
     */
    function removeCitation(fieldId, activeStyle) {
        var isBookmark = (fieldId && String(fieldId).indexOf("ZOTERO_BREF_") === 0);
        if (isBookmark) {
            window.Asc.scope.delBmark = fieldId;
            return new Promise(function (resolve) {
                window.Asc.plugin.callCommand(function () {
                    var doc = Api.GetDocument();
                    if (doc.DeleteBookmark) doc.DeleteBookmark(Asc.scope.delBmark);
                    var props = doc.GetCustomProperties ? doc.GetCustomProperties() : null;
                    if (props && props.Remove) props.Remove(Asc.scope.delBmark);
                }, false, true, function () {
                    autoUpgradeBibliography(activeStyle).then(resolve).catch(resolve);
                });
            });
        }

        return new Promise(function (resolve) {
            try {
                window.Asc.plugin.executeMethod("RemoveAddinField", [fieldId]);
            } catch (e) {
                console.warn("RemoveAddinField error:", e);
            }
            autoUpgradeBibliography(activeStyle).then(resolve).catch(resolve);
        });
    }

    /**
     * Decode HTML entities and clean Unicode directional isolates without stripping inter-tag spaces
     */
    function cleanNodeText(str) {
        if (!str) return "";
        return str
            .replace(/&#38;/g, "&")
            .replace(/&amp;/g, "&")
            .replace(/&#60;/g, "<")
            .replace(/&lt;/g, "<")
            .replace(/&#62;/g, ">")
            .replace(/&gt;/g, ">")
            .replace(/[\u2066-\u206f\u200b-\u200f]/g, "")
            .replace(/\r?\n|\r/g, " ")
            .replace(/[ \t]+/g, " ");
    }

    /**
     * Parse HTML bibliography entries into text runs with formatting (italics, bold)
     * Preserves interstitial spaces between plain text and formatted elements (e.g. "Review. <i>RELC Journal</i>, <i>57</i>").
     */
    function parseBibHtml(html) {
        var div = document.createElement("div");
        div.innerHTML = html;

        var entryNodes = div.querySelectorAll(".csl-entry");
        if (!entryNodes.length) {
            entryNodes = div.querySelectorAll("p");
        }
        if (!entryNodes.length) {
            entryNodes = [div];
        }

        var parsedEntries = [];
        entryNodes.forEach(function (entryNode) {
            var runs = [];

            function walk(node, isItalic, isBold) {
                if (node.nodeType === Node.TEXT_NODE) {
                    var text = cleanNodeText(node.textContent);
                    if (text && text.length > 0) {
                        runs.push({ text: text, italic: !!isItalic, bold: !!isBold });
                    }
                } else if (node.nodeType === Node.ELEMENT_NODE) {
                    var tag = node.tagName.toLowerCase();
                    var curItalic = isItalic || (tag === "i" || tag === "em");
                    var curBold = isBold || (tag === "b" || tag === "strong");

                    for (var i = 0; i < node.childNodes.length; i++) {
                        walk(node.childNodes[i], curItalic, curBold);
                    }
                }
            }

            walk(entryNode, false, false);
            if (runs.length > 0) {
                // Clean boundary spaces for the overall entry without collapsing inter-run spacing
                runs[0].text = runs[0].text.replace(/^\s+/, "");
                runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, "");
                runs = runs.filter(function (r) { return r.text.length > 0; });
                if (runs.length > 0) {
                    parsedEntries.push(runs);
                }
            }
        });

        return parsedEntries;
    }

    /**
     * Clean and prepare HTML string for insertion via PasteHtml
     */
    function prepareBibHtmlForPaste(html) {
        var div = document.createElement("div");
        div.innerHTML = html;

        var entries = div.querySelectorAll(".csl-entry");
        if (entries && entries.length > 0) {
            var output = "";
            entries.forEach(function (e) {
                // Ensure each bibliography entry is wrapped in a paragraph tag for Word formatting
                output += "<p style='font-family:\"Times New Roman\", serif; margin-left: 36pt; text-indent: -36pt; line-height: 1.15; margin-bottom: 7pt;'>" + e.innerHTML + "</p>";
            });
            return output;
        }
        return html;
    }

    /**
     * Apply hanging indent (720 / -720) and Times New Roman font to selected paragraphs
     */
    function applyBibParagraphFormatting() {
        return new Promise(function (resolve) {
            window.Asc.plugin.callCommand(function () {
                var doc = Api.GetDocument();
                var range = doc.GetRangeBySelect();
                if (range && range.GetAllParagraphs) {
                    var paras = range.GetAllParagraphs();
                    if (paras && paras.length) {
                        for (var i = 0; i < paras.length; i++) {
                            var p = paras[i];
                            p.SetFontFamily("Times New Roman");
                            p.SetIndLeft(720);        // 0.5" / 1.27 cm Left Indent
                            p.SetIndFirstLine(-720);  // 悬挂缩进 (Hanging Indent)
                            p.SetSpacingAfter(140);   // ~7pt space
                            p.SetSpacingLine(276, "multiple");
                        }
                    }
                }
            }, false, true, function () {
                resolve();
            });
        });
    }

    /**
     * Insert or update bibliography in document with standard Hanging Indent (悬挂缩进)
     * Every single entry is created as its own ApiParagraph with 0.5" hanging indent,
     * Times New Roman, and full CSL bold/italic runs.
     */
    function updateBibliographyInDoc(htmlBib, entries) {
        return new Promise(function (resolve, reject) {
            // Sort entries alphabetically (A -> Z)
            if (Array.isArray(entries)) {
                entries.sort(function (a, b) {
                    var textA = a.map(function (r) { return r.text; }).join("").trim().toLowerCase();
                    var textB = b.map(function (r) { return r.text; }).join("").trim().toLowerCase();
                    return textA.localeCompare(textB);
                });
            }

            window.Asc.scope.bibEntries = entries;

            // Remove any legacy Word AddinField for the bibliography
            window.Asc.plugin.executeMethod("GetAllAddinFields", null, function (fields) {
                var oldBibFields = [];
                if (Array.isArray(fields)) {
                    oldBibFields = fields.filter(function (f) {
                        return f && f.Value && (f.Value.indexOf("ZOTERO_BIBL") !== -1 || f.Value.indexOf("ZOTERO_BIBLIOGRAPHY") !== -1);
                    });
                }

                function proceedWithDocUpdate() {
                    window.Asc.plugin.callCommand(function () {
                        var doc = Api.GetDocument();
                        var entries = Asc.scope.bibEntries;

                        // 1. Check if ContentControl already exists:
                        var controls = doc.GetContentControlsByTag("ZOTERO_BIBLIOGRAPHY");
                        if (controls && controls.length > 0) {
                            var oBlockLvlSdt = controls[0];
                            oBlockLvlSdt.RemoveAllElements();

                            for (var i = 0; i < entries.length; i++) {
                                var oParagraph = Api.CreateParagraph();
                                oParagraph.SetFontFamily("Times New Roman");
                                oParagraph.SetIndLeft(720);        // 0.5" hanging indent
                                oParagraph.SetIndFirstLine(-720);  // -0.5" first line
                                oParagraph.SetSpacingAfter(140);   // ~7pt space
                                oParagraph.SetSpacingLine(276, "multiple");

                                var runs = entries[i];
                                for (var j = 0; j < runs.length; j++) {
                                    var rData = runs[j];
                                    var oRun = oParagraph.AddText(rData.text);
                                    oRun.SetFontFamily("Times New Roman");
                                    if (rData.italic) oRun.SetItalic(true);
                                    if (rData.bold) oRun.SetBold(true);
                                }
                                if (oBlockLvlSdt.Push) {
                                    oBlockLvlSdt.Push(oParagraph);
                                } else {
                                    oBlockLvlSdt.AddElement(oParagraph, i);
                                }
                            }
                            return true;
                        }

                        // 2. If ContentControl does not exist (e.g. Manu.docx from Word):
                        // Look for the "References" or "Bibliography" heading paragraph:
                        var paras = doc.GetAllParagraphs();
                        var refPara = null;
                        var refIdx = -1;

                        for (var p = 0; p < paras.length; p++) {
                            var text = paras[p].GetText().trim();
                            if (/^(references|bibliography|参考文献)\b/i.test(text)) {
                                refPara = paras[p];
                                refIdx = p;
                                break;
                            }
                        }

                        // Create the clean BlockLvlSdt container
                        var oNewBlockLvlSdt = Api.CreateBlockLvlSdt();
                        oNewBlockLvlSdt.SetTag("ZOTERO_BIBLIOGRAPHY");
                        oNewBlockLvlSdt.SetAlias("Zotero Bibliography");

                        for (var e = 0; e < entries.length; e++) {
                            var oPara = Api.CreateParagraph();
                            oPara.SetFontFamily("Times New Roman");
                            oPara.SetIndLeft(720);        // 0.5" hanging indent
                            oPara.SetIndFirstLine(-720);  // -0.5" first line
                            oPara.SetSpacingAfter(140);   // ~7pt space
                            oPara.SetSpacingLine(276, "multiple");

                            var rList = entries[e];
                            for (var k = 0; k < rList.length; k++) {
                                var runData = rList[k];
                                var rObj = oPara.AddText(runData.text);
                                rObj.SetFontFamily("Times New Roman");
                                if (runData.italic) rObj.SetItalic(true);
                                if (runData.bold) rObj.SetBold(true);
                            }
                            if (oNewBlockLvlSdt.Push) {
                                oNewBlockLvlSdt.Push(oPara);
                            } else {
                                oNewBlockLvlSdt.AddElement(oPara, e);
                            }
                        }

                        if (refPara && refIdx !== -1) {
                            // Delete any old trailing bibliography paragraphs following References:
                            for (var d = paras.length - 1; d > refIdx; d--) {
                                paras[d].Delete();
                            }
                            // Move cursor to end (immediately after References) and insert the new bibliography:
                            doc.MoveCursorToEnd();
                            doc.InsertContent([oNewBlockLvlSdt]);
                        } else {
                            // Brand new document or cursor location:
                            doc.InsertContent([oNewBlockLvlSdt]);
                        }

                        return true;
                    }, false, true, function () {
                        resolve(entries ? entries.length : 1);
                    });
                }

                if (oldBibFields.length > 0) {
                    oldBibFields.forEach(function (f) {
                        try {
                            window.Asc.plugin.executeMethod("RemoveAddinField", [f.FieldId]);
                        } catch (e) {}
                    });
                }
                proceedWithDocUpdate();
            });
        });
    }

    /**
     * Insert bibliography into document with standard Hanging Indent (悬挂缩进)
     * and strictly defaults to Times New Roman.
     */
    function insertBibliography(style) {
        return getAllCitationFields().then(function (fields) {
            if (!fields.length) {
                throw new Error("No citations found in document. Insert at least one citation first.");
            }

            var keysSet = {};
            fields.forEach(function (f) {
                (f.citationItems || []).forEach(function (it) {
                    var k = getItemKey(it);
                    if (k) keysSet[k] = true;
                });
            });
            var keys = Object.keys(keysSet);
            if (!keys.length) {
                throw new Error("No valid citation keys found.");
            }

            return ZoteroAPI.getBibliography(keys, 0, style).then(function (htmlBib) {
                if (!htmlBib) {
                    throw new Error("Failed to generate bibliography from Zotero.");
                }

                var entries = parseBibHtml(htmlBib);
                if (!entries.length) {
                    throw new Error("No bibliography entries could be parsed.");
                }

                setDocumentZoteroStyle(style);
                return updateBibliographyInDoc(htmlBib, entries);
            });
        });
    }

    /**
     * Automatically upgrade the bibliography list in the document whenever citations change
     */
    function autoUpgradeBibliography(style, extraKeys) {
        return new Promise(function (resolve) {
            var finished = false;
            var timer = setTimeout(function () {
                if (!finished) {
                    finished = true;
                    resolve();
                }
            }, 3000);

            getBibliographyInfo().then(function (bibInfo) {
                if (!bibInfo) {
                    if (!finished) {
                        finished = true;
                        clearTimeout(timer);
                        resolve();
                    }
                    return;
                }

                getAllCitationFields().then(function (fields) {
                    var keysSet = {};
                    fields.forEach(function (f) {
                        (f.citationItems || []).forEach(function (it) {
                            var k = getItemKey(it);
                            if (k) keysSet[k] = true;
                        });
                    });

                    if (Array.isArray(extraKeys)) {
                        extraKeys.forEach(function (k) {
                            if (k) keysSet[k] = true;
                        });
                    }

                    var keys = Object.keys(keysSet);
                    if (!keys.length) {
                        if (!finished) {
                            finished = true;
                            clearTimeout(timer);
                            resolve();
                        }
                        return;
                    }

                    ZoteroAPI.getBibliography(keys, 0, style).then(function (htmlBib) {
                        if (!htmlBib) {
                            if (!finished) {
                                finished = true;
                                clearTimeout(timer);
                                resolve();
                            }
                            return;
                        }
                        var entries = parseBibHtml(htmlBib);
                        if (!entries.length) {
                            if (!finished) {
                                finished = true;
                                clearTimeout(timer);
                                resolve();
                            }
                            return;
                        }

                        updateBibliographyInDoc(htmlBib, entries).then(function () {
                            if (!finished) {
                                finished = true;
                                clearTimeout(timer);
                                resolve();
                            }
                        }).catch(function () {
                            if (!finished) {
                                finished = true;
                                clearTimeout(timer);
                                resolve();
                            }
                        });
                    }).catch(function () {
                        if (!finished) {
                            finished = true;
                            clearTimeout(timer);
                            resolve();
                        }
                    });
                }).catch(function () {
                    if (!finished) {
                        finished = true;
                        clearTimeout(timer);
                        resolve();
                    }
                });
            }).catch(function () {
                if (!finished) {
                    finished = true;
                    clearTimeout(timer);
                    resolve();
                }
            });
        });
    }

    /**
     * Reformat all citations and bibliography across the entire document to a new CSL style.
     * Concurrently retrieves items and bibliography from Zotero to cut reformat latency in half!
     */
    function reformatDocumentStyle(newStyle) {
        return getAllCitationFields().then(function (fields) {
            if (!fields.length) {
                return autoUpgradeBibliography(newStyle).then(function () {
                    return 0;
                });
            }

            var allKeysSet = {};
            fields.forEach(function (f) {
                (f.citationItems || []).forEach(function (it) {
                    var k = getItemKey(it);
                    if (k) allKeysSet[k] = true;
                });
            });
            var allKeys = Object.keys(allKeysSet);
            if (!allKeys.length) return 0;

            // CONCURRENT PIPELINE: Fetch in-text items and bibliography concurrently
            return Promise.all([
                ZoteroAPI.getItemsByKeys(allKeys, 0, newStyle),
                ZoteroAPI.getBibliography(allKeys, 0, newStyle)
            ]).then(function (results) {
                var items = results[0] || [];
                var htmlBib = results[1] || "";

                var itemMap = {};
                items.forEach(function (it) {
                    itemMap[it.key] = it;
                });

                // Fallback to itemData if any item was not returned by local Zotero
                fields.forEach(function (f) {
                    (f.citationItems || []).forEach(function (citItem) {
                        var k = getItemKey(citItem);
                        if (k && !itemMap[k]) {
                            itemMap[k] = createItemFromItemData(citItem, k);
                        }
                    });
                });

                var updatedFields = [];
                fields.forEach(function (f) {
                    var clusterItems = (f.citationItems || []).map(function (it) {
                        var k = getItemKey(it);
                        return itemMap[k] || createItemFromItemData(it, k);
                    });

                    var firstCit = (f.citationItems && f.citationItems[0]) || {};
                    var isNarrative = !!(f.properties && f.properties.isNarrative) || !!firstCit.isNarrative;
                    var isSuppress = !!firstCit["suppress-author"] && !isNarrative;

                    var formatMode = "parenthetical";
                    if (isNarrative) {
                        formatMode = "narrative";
                    } else if (isSuppress) {
                        formatMode = "suppress";
                    }

                    var built = buildCitation(clusterItems, {
                        locator: firstCit.locator || "",
                        prefix: firstCit.prefix || "",
                        suffix: firstCit.suffix || "",
                        formatMode: formatMode,
                        style: newStyle,
                        citationID: f.citationID,
                        isFootnote: f.properties ? f.properties.noteIndex > 0 : false
                    });

                    updatedFields.push({
                        FieldId: f.fieldId,
                        Value: CIT_PREFIX + " " + JSON.stringify(built.citationObj),
                        Content: built.formattedText
                    });
                });

                var pDocFields = updatedFields.length > 0 ? new Promise(function (res) {
                    window.Asc.plugin.executeMethod("UpdateAddinFields", [updatedFields], res);
                }) : Promise.resolve();

                var entries = parseBibHtml(htmlBib);
                var pBib = entries.length > 0 ? updateBibliographyInDoc(htmlBib, entries) : Promise.resolve();

                return Promise.all([pDocFields, pBib]).then(function () {
                    setDocumentZoteroStyle(newStyle);
                    return updatedFields.length;
                });
            });
        });
    }

    /**
     * Unlink all citations (convert fields and bookmarks to plain text)
     */
    function unlinkAll() {
        return new Promise(function (resolve) {
            window.Asc.plugin.executeMethod("GetAllAddinFields", null, function (fields) {
                var zoteroFields = [];
                if (Array.isArray(fields)) {
                    zoteroFields = fields.filter(function (f) {
                        return f && f.Value && (f.Value.indexOf("ZOTERO_ITEM") !== -1 || f.Value.indexOf("ZOTERO_BIBL") !== -1 || f.Value.indexOf("ZOTERO_BIBLIOGRAPHY") !== -1);
                    });
                }

                function doCleanBookmarksAndControls(addinsUnlinked) {
                    window.Asc.plugin.callCommand(function () {
                        var doc = Api.GetDocument();
                        var count = 0;
                        var bmarks = doc.GetAllBookmarksNames ? doc.GetAllBookmarksNames() : [];
                        for (var i = 0; i < bmarks.length; i++) {
                            if (bmarks[i].indexOf("ZOTERO_BREF_") === 0) {
                                if (doc.DeleteBookmark) doc.DeleteBookmark(bmarks[i]);
                                count++;
                            }
                        }
                        return count;
                    }, false, true, function (bmarkCount) {
                        resolve((addinsUnlinked || 0) + (bmarkCount || 0));
                    });
                }

                if (!zoteroFields.length) {
                    doCleanBookmarksAndControls(0);
                    return;
                }

                var remaining = zoteroFields.length;
                zoteroFields.forEach(function (f) {
                    window.Asc.plugin.executeMethod("RemoveFieldWrapper", [f.FieldId], function () {
                        remaining--;
                        if (remaining <= 0) {
                            doCleanBookmarksAndControls(zoteroFields.length);
                        }
                    });
                });
            });
        });
    }

    return {
        getItemKey: getItemKey,
        createItemFromItemData: createItemFromItemData,
        parseCitationField: parseCitationField,
        getCurrentField: getCurrentField,
        getAllCitationFields: getAllCitationFields,
        getBibliographyInfo: getBibliographyInfo,
        hasBibliography: hasBibliography,
        hasBibliographyControl: hasBibliography, // backwards compatibility
        getDocumentZoteroStyle: getDocumentZoteroStyle,
        setDocumentZoteroStyle: setDocumentZoteroStyle,
        buildCitation: buildCitation,
        extractNarrativeFromZoteroStyle: extractNarrativeFromZoteroStyle,
        insertCitation: insertCitation,
        updateCitation: updateCitation,
        removeCitation: removeCitation,
        insertBibliography: insertBibliography,
        autoUpgradeBibliography: autoUpgradeBibliography,
        reformatDocumentStyle: reformatDocumentStyle,
        unlinkAll: unlinkAll
    };
})();
