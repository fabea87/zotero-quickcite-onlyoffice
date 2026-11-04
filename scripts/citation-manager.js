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
     * Get the addin field at current cursor position
     */
    function getCurrentField() {
        return new Promise(function (resolve) {
            try {
                window.Asc.plugin.executeMethod("GetCurrentAddinField", null, function (field) {
                    if (field && field.Value && field.Value.indexOf("ZOTERO_ITEM") !== -1) {
                        resolve(parseCitationField(field));
                    } else {
                        resolve(null);
                    }
                });
            } catch (e) {
                console.warn("GetCurrentAddinField failed:", e);
                resolve(null);
            }
        });
    }

    /**
     * Get all Zotero citation fields in document
     */
    function getAllCitationFields() {
        return new Promise(function (resolve) {
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
                    resolve(zoteroFields);
                });
            } catch (e) {
                console.warn("GetAllAddinFields failed:", e);
                resolve([]);
            }
        });
    }

    /**
     * Locate existing bibliography representation in document:
     * Checks Word AddinField (ADDIN ZOTERO_BIBL / ZOTERO_BIBLIOGRAPHY) and ONLYOFFICE ContentControl (ZOTERO_BIBLIOGRAPHY)
     */
    function getBibliographyInfo() {
        return new Promise(function (resolve) {
            try {
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
                        resolve(null);
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
     * Build formatted citation text and Word CSL data structure.
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

        var citationItems = items.map(function (item, idx) {
            var isPrimary = (idx === 0);
            var itemLocator = isPrimary ? locator : (item.locator || "");
            var itemPrefix = isPrimary ? prefix : (item.prefix || "");
            var itemSuffix = isPrimary ? suffix : (item.suffix || "");
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
            var yearParts = items.map(function (item, idx) {
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
            var parts = items.map(function (item, idx) {
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
     * Insert citation into document, strictly setting Times New Roman font,
     * syncing document style, and kicking off autoUpgradeBibliography in the background
     */
    function insertCitation(citationData, isFootnote, activeStyle) {
        var fieldId = "ZOTERO_" + Date.now();
        var addinField = {
            FieldId: fieldId,
            Value: CIT_PREFIX + " " + JSON.stringify(citationData.citationObj),
            Content: citationData.formattedText
        };

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

        return doAddField().then(function (result) {
            setDocumentZoteroStyle(activeStyle);
            var newKeys = (citationData.citationObj.citationItems || []).map(function (it) {
                return getItemKey(it);
            });
            setTimeout(function () {
                autoUpgradeBibliography(activeStyle, newKeys);
            }, 50);
            return result;
        });
    }

    /**
     * Update existing citation in document, then awaits auto-upgrading bibliography
     */
    function updateCitation(fieldId, citationData, activeStyle) {
        var updatedField = {
            FieldId: fieldId,
            Value: CIT_PREFIX + " " + JSON.stringify(citationData.citationObj),
            Content: citationData.formattedText
        };

        return new Promise(function (resolve) {
            window.Asc.plugin.executeMethod("UpdateAddinFields", [[updatedField]], function () {
                applyFontToField(fieldId).then(function () {
                    setDocumentZoteroStyle(activeStyle);
                    var extraKeys = (citationData.citationObj.citationItems || []).map(function (it) {
                        return getItemKey(it);
                    });
                    setTimeout(function () {
                        autoUpgradeBibliography(activeStyle, extraKeys);
                    }, 50);
                    resolve(updatedField);
                }).catch(function () {
                    resolve(updatedField);
                });
            });
        });
    }

    /**
     * Remove citation field from document, then awaits auto-upgrading bibliography
     */
    function removeCitation(fieldId, activeStyle) {
        return new Promise(function (resolve) {
            window.Asc.plugin.executeMethod("RemoveAddinField", [fieldId], function () {
                setTimeout(function () {
                    autoUpgradeBibliography(activeStyle);
                }, 50);
                resolve();
            });
        });
    }

    /**
     * Parse HTML bibliography entries into text runs with formatting (italics, bold)
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
                    var text = cleanText(node.textContent);
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
                parsedEntries.push(runs);
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
     * Supports both MS Word-compatible AddinField (ADDIN ZOTERO_BIBL) and ONLYOFFICE ContentControl.
     */
    function updateBibliographyInDoc(htmlBib, entries, bibInfo) {
        return new Promise(function (resolve, reject) {
            var bibLines = entries.map(function (e) {
                return e.map(function (r) { return r.text; }).join("");
            });
            var allBibText = bibLines.join("\n\n");

            // 1. If an AddinField exists in document (Word format from Manu.docx or previous save):
            if (bibInfo && bibInfo.type === "addinField" && bibInfo.field && bibInfo.field.FieldId) {
                var updatedBibField = {
                    FieldId: bibInfo.field.FieldId,
                    Value: bibInfo.field.Value || BIB_PREFIX,
                    Content: allBibText
                };

                var finished = false;
                var timeout = setTimeout(function () {
                    if (!finished) {
                        finished = true;
                        resolve(entries.length);
                    }
                }, 1200);

                try {
                    window.Asc.plugin.executeMethod("UpdateAddinFields", [[updatedBibField]], function () {
                        window.Asc.scope.bibEntries = entries;
                        window.Asc.plugin.callCommand(function () {
                            var doc = Api.GetDocument();
                            var paras = doc.GetAllParagraphs();
                            var inRef = false;
                            for (var p = 0; p < paras.length; p++) {
                                var text = paras[p].GetText().trim();
                                if (!inRef) {
                                    if (/^(references|bibliography|参考文献)\b/i.test(text)) {
                                        inRef = true;
                                    }
                                } else {
                                    paras[p].SetFontFamily("Times New Roman");
                                    paras[p].SetIndLeft(720);        // 0.5" hanging indent
                                    paras[p].SetIndFirstLine(-720);  // -0.5" first line
                                    paras[p].SetSpacingAfter(140);
                                    paras[p].SetSpacingLine(276, "multiple");
                                }
                            }
                        }, false, true, function () {
                            if (!finished) {
                                finished = true;
                                clearTimeout(timeout);
                                resolve(entries.length);
                            }
                        });
                    });
                } catch (e) {
                    if (!finished) {
                        finished = true;
                        clearTimeout(timeout);
                        resolve(entries.length);
                    }
                }
                return;
            }

            // 2. If a ContentControl exists (ONLYOFFICE BlockLvlSdt format):
            if (bibInfo && bibInfo.type === "contentControl") {
                window.Asc.scope.bibEntries = entries;
                window.Asc.plugin.callCommand(function () {
                    var oDocument = Api.GetDocument();
                    var entries = Asc.scope.bibEntries;
                    var controls = oDocument.GetContentControlsByTag("ZOTERO_BIBLIOGRAPHY");
                    if (!controls || controls.length === 0) return;
                    var oBlockLvlSdt = controls[0];
                    oBlockLvlSdt.RemoveAllElements();
                    for (var i = 0; i < entries.length; i++) {
                        var oParagraph = Api.CreateParagraph();
                        oParagraph.SetFontFamily("Times New Roman");
                        oParagraph.SetIndLeft(720);
                        oParagraph.SetIndFirstLine(-720);
                        oParagraph.SetSpacingAfter(140);
                        oParagraph.SetSpacingLine(276, "multiple");
                        var runs = entries[i];
                        for (var j = 0; j < runs.length; j++) {
                            var rData = runs[j];
                            var oRun = oParagraph.AddText(rData.text);
                            oRun.SetFontFamily("Times New Roman");
                            if (rData.italic) oRun.SetItalic(true);
                            if (rData.bold) oRun.SetBold(true);
                        }
                        oBlockLvlSdt.AddElement(oParagraph);
                    }
                }, false, true, function () {
                    resolve(entries.length);
                });
                return;
            }

            // 3. If neither exists (brand new document insertion):
            window.Asc.scope.bibEntries = entries;
            window.Asc.plugin.callCommand(function () {
                var oDocument = Api.GetDocument();
                var entries = Asc.scope.bibEntries;
                var oBlockLvlSdt = Api.CreateBlockLvlSdt();
                oBlockLvlSdt.SetTag("ZOTERO_BIBLIOGRAPHY");
                oBlockLvlSdt.SetAlias("Zotero Bibliography");

                for (var i = 0; i < entries.length; i++) {
                    var oParagraph = Api.CreateParagraph();
                    oParagraph.SetFontFamily("Times New Roman");
                    oParagraph.SetIndLeft(720);
                    oParagraph.SetIndFirstLine(-720);
                    oParagraph.SetSpacingAfter(140);
                    oParagraph.SetSpacingLine(276, "multiple");

                    var runs = entries[i];
                    for (var j = 0; j < runs.length; j++) {
                        var rData = runs[j];
                        var oRun = oParagraph.AddText(rData.text);
                        oRun.SetFontFamily("Times New Roman");
                        if (rData.italic) oRun.SetItalic(true);
                        if (rData.bold) oRun.SetBold(true);
                    }
                    oBlockLvlSdt.AddElement(oParagraph);
                }
                oDocument.InsertContent([oBlockLvlSdt]);
            }, false, true, function () {
                resolve(entries ? entries.length : 1);
            });
        });
    }

    /**
     * Insert bibliography into document with standard Hanging Indent (悬挂缩进)
     * and strictly defaults to Times New Roman.
     */
    function insertBibliography(style) {
        return getBibliographyInfo().then(function (existingInfo) {
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
                    return updateBibliographyInDoc(htmlBib, entries, existingInfo);
                });
            });
        });
    }

    /**
     * Automatically upgrade the bibliography list in the document whenever citations change
     */
    function autoUpgradeBibliography(style, extraKeys) {
        return getBibliographyInfo().then(function (bibInfo) {
            if (!bibInfo) return; // Only upgrade if a bibliography already exists in the document

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
                if (!keys.length) return;

                return ZoteroAPI.getBibliography(keys, 0, style).then(function (htmlBib) {
                    if (!htmlBib) return;
                    var entries = parseBibHtml(htmlBib);
                    if (!entries.length) return;

                    return updateBibliographyInDoc(htmlBib, entries, bibInfo);
                });
            });
        }).catch(function (err) {
            console.warn("Auto-upgrade error:", err);
        });
    }

    /**
     * Reformat all citations and bibliography across the entire document to a new CSL style.
     * Fully compatible with citations from MS Word or ONLYOFFICE.
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

            return ZoteroAPI.getItemsByKeys(allKeys, 0, newStyle).then(function (items) {
                var itemMap = {};
                (items || []).forEach(function (it) {
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

                if (updatedFields.length > 0) {
                    return new Promise(function (resolve) {
                        window.Asc.plugin.executeMethod("UpdateAddinFields", [updatedFields], function () {
                            setDocumentZoteroStyle(newStyle);
                            autoUpgradeBibliography(newStyle, allKeys).then(function () {
                                resolve(updatedFields.length);
                            }).catch(function () {
                                resolve(updatedFields.length);
                            });
                        });
                    });
                }
                return 0;
            });
        });
    }

    /**
     * Unlink all citations (convert fields to plain text)
     */
    function unlinkAll() {
        return new Promise(function (resolve) {
            window.Asc.plugin.executeMethod("GetAllAddinFields", null, function (fields) {
                if (!Array.isArray(fields) || !fields.length) {
                    resolve(0);
                    return;
                }
                var zoteroFields = fields.filter(function (f) {
                    return f.Value && (f.Value.indexOf("ZOTERO_ITEM") !== -1 || f.Value.indexOf("ZOTERO_BIBL") !== -1 || f.Value.indexOf("ZOTERO_BIBLIOGRAPHY") !== -1);
                });
                var count = zoteroFields.length;
                if (!count) {
                    resolve(0);
                    return;
                }
                var remaining = count;
                zoteroFields.forEach(function (f) {
                    window.Asc.plugin.executeMethod("RemoveFieldWrapper", [f.FieldId], function () {
                        remaining--;
                        if (remaining <= 0) resolve(count);
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
