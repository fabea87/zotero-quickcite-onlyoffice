/**
 * Zotero QuickCite - Citation & Document Manager
 * Manages Word-compatible AddinFields, CSL formatting, hanging indents (悬挂缩进),
 * ubiquitous narrative citations (with multi-language author switching like 'Yan et al. (2026)' <-> '严等 (2026)'),
 * automatic bibliography upgrades, and document-wide style changes in Times New Roman.
 */
var CitationManager = (function () {
    "use strict";

    var CIT_PREFIX = "ZOTERO_ITEM CSL_CITATION";

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
                                zoteroFields.push(parseCitationField(f));
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
     * Check if a bibliography ContentControl exists in document
     */
    function hasBibliographyControl() {
        return new Promise(function (resolve) {
            try {
                window.Asc.plugin.executeMethod("GetAllContentControls", null, function (controls) {
                    if (Array.isArray(controls) && controls.length > 0) {
                        var found = controls.some(function (c) {
                            return c && (c.Tag === "ZOTERO_BIBLIOGRAPHY" || c.InternalId === "ZOTERO_BIBLIOGRAPHY");
                        });
                        if (found) {
                            resolve(true);
                            return;
                        }
                    }
                    resolve(false);
                });
            } catch (e) {
                resolve(false);
            }
        });
    }

    /**
     * Parse raw field into structured citation object
     */
    function parseCitationField(rawField) {
        var jsonStr = rawField.Value.replace(/^ZOTERO_ITEM\s+CSL_CITATION\s*/, "").trim();
        var data = {};
        try {
            data = JSON.parse(jsonStr);
        } catch (e) {
            console.warn("Could not parse citation JSON:", jsonStr);
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
     * Fallback prose author calculation if Zotero API citation is missing
     */
    function computeNarrativeAuthor(item, isChineseStyle) {
        if (!item) return "";
        var rawData = item.raw ? (item.raw.data || {}) : {};
        var creators = rawData.creators || [];

        if (Array.isArray(creators) && creators.length > 0) {
            var first = creators[0];
            var firstLast = cleanText(first.lastName || first.name || "");

            if (creators.length === 1) {
                return firstLast;
            } else if (creators.length === 2) {
                var second = creators[1];
                var secondLast = cleanText(second.lastName || second.name || "");
                return isChineseStyle ? (firstLast + "和" + secondLast) : (firstLast + " and " + secondLast);
            } else {
                return isChineseStyle ? (firstLast + "等") : (firstLast + " et al.");
            }
        }

        var authors = cleanText(item.authors || "Author");
        return isChineseStyle ? authors.replace(/\s*&\s*/g, "和").replace(/et al\./g, "等") : authors.replace(/\s*&\s*/g, " and ").replace(/\s+/g, " ");
    }

    /**
     * Ubiquitously extract narrative author and year dynamically from Zotero's CSL output.
     * Works across ALL styles and ALL languages (English, Chinese, German, French, etc.):
     * - English: "(Wang et al., 2026)" -> "Wang et al. (2026)"
     * - Chinese: "(Wang等, 2026)" -> "Wang等 (2026)" or "(严等, 2026)" -> "严等 (2026)"
     * - German: "(Schmidt u. a., 2024)" -> "Schmidt u. a. (2024)"
     * - French: "(Dupont et coll., 2023)" -> "Dupont et coll. (2023)"
     * - 2 authors: "(Teng & Zhang, 2016)" -> "Teng and Zhang (2016)"
     */
    function universalExtractNarrative(zoteroCit, year, locator, prefix, suffix, fallbackItem) {
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
        } else if (fallbackItem) {
            authorPart = computeNarrativeAuthor(fallbackItem, /[\u4e00-\u9fa5]/.test(raw));
        } else {
            authorPart = inner;
        }

        // In narrative prose, if author string has '&', replace with language conjunction
        if (authorPart.indexOf("&") !== -1) {
            if (/[\u4e00-\u9fa5]/.test(authorPart)) {
                authorPart = authorPart.replace(/\s*&\s*/g, "和");
            } else {
                authorPart = authorPart.replace(/\s*&\s*/g, " and ");
            }
        }

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
     * In narrative mode, the entire "Author (Year)" string is unified inside the AddinField,
     * ensuring automatic bilingual transformation (e.g. 'Yan et al. (2026)' <-> '严等 (2026)') when style changes!
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

            return {
                id: item.key || item.id,
                uris: ["http://zotero.org/users/0/items/" + (item.key || item.id)],
                itemData: item.raw ? (item.raw.data || {}) : {},
                locator: itemLocator,
                label: itemLocator ? "page" : undefined,
                prefix: itemPrefix,
                suffix: itemSuffix,
                "suppress-author": itemSuppress,
                isNarrative: isNarrative
            };
        });

        if (isNarrative && items.length > 0) {
            // NARRATIVE CITATION: e.g. "Yan et al. (2026)" or "Yan等 (2026)"
            var primaryItem = items[0];
            var narrative = universalExtractNarrative(
                primaryItem.citation,
                primaryItem.year,
                locator,
                prefix,
                suffix,
                primaryItem
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
            window.Asc.plugin.executeMethod("SelectAddinField", [fieldId], function () {
                window.Asc.plugin.callCommand(function () {
                    var oDocument = Api.GetDocument();
                    var oRange = oDocument.GetRangeBySelect();
                    if (oRange && oRange.SetFontFamily) {
                        oRange.SetFontFamily("Times New Roman");
                    }
                }, false, true, function () {
                    window.Asc.plugin.executeMethod("MoveCursorOutsideField", [fieldId, false], function () {
                        resolve();
                    });
                });
            });
        });
    }

    /**
     * Insert citation into document, strictly setting Times New Roman font,
     * and awaiting autoUpgradeBibliography
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
                            });
                        });
                    });
                });
            }
        }

        return doAddField().then(function (result) {
            var newKeys = (citationData.citationObj.citationItems || []).map(function (it) {
                return it.id;
            });
            return autoUpgradeBibliography(activeStyle, newKeys).then(function () {
                return result;
            });
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
                    var extraKeys = (citationData.citationObj.citationItems || []).map(function (it) {
                        return it.id;
                    });
                    autoUpgradeBibliography(activeStyle, extraKeys).then(function () {
                        resolve(updatedField);
                    });
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
                autoUpgradeBibliography(activeStyle).then(resolve);
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
                    if (it.id) keysSet[it.id] = true;
                });
            });
            var keys = Object.keys(keysSet);
            if (!keys.length) {
                throw new Error("No valid citation keys found.");
            }

            return ZoteroAPI.getBibliography(keys, 0, style);
        }).then(function (htmlBib) {
            if (!htmlBib) {
                throw new Error("Failed to generate bibliography from Zotero.");
            }

            var entries = parseBibHtml(htmlBib);
            if (!entries.length) {
                throw new Error("No bibliography entries could be parsed.");
            }

            return new Promise(function (resolve) {
                window.Asc.scope.bibEntries = entries;

                window.Asc.plugin.callCommand(function () {
                    var oDocument = Api.GetDocument();
                    var entries = Asc.scope.bibEntries;

                    var oBlockLvlSdt = null;
                    var controls = oDocument.GetContentControlsByTag("ZOTERO_BIBLIOGRAPHY");

                    if (controls && controls.length > 0) {
                        oBlockLvlSdt = controls[0];
                        oBlockLvlSdt.RemoveAllElements();
                    } else {
                        oBlockLvlSdt = Api.CreateBlockLvlSdt();
                        oBlockLvlSdt.SetTag("ZOTERO_BIBLIOGRAPHY");
                        oBlockLvlSdt.SetAlias("Zotero Bibliography");
                    }

                    for (var i = 0; i < entries.length; i++) {
                        var oParagraph = Api.CreateParagraph();

                        oParagraph.SetFontFamily("Times New Roman");
                        oParagraph.SetIndLeft(720);        // 0.5" / 1.27 cm Left Indent
                        oParagraph.SetIndFirstLine(-720);  // 悬挂缩进 (Hanging Indent)
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

                        oBlockLvlSdt.AddElement(oParagraph);
                    }

                    if (!controls || controls.length === 0) {
                        oDocument.InsertContent([oBlockLvlSdt]);
                    }
                }, false, true, function () {
                    resolve(entries.length);
                });
            });
        });
    }

    /**
     * Automatically upgrade the bibliography list in the document whenever citations change
     */
    function autoUpgradeBibliography(style, extraKeys) {
        return getAllCitationFields().then(function (fields) {
            var keysSet = {};
            fields.forEach(function (f) {
                (f.citationItems || []).forEach(function (it) {
                    if (it.id) keysSet[it.id] = true;
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

                return new Promise(function (resolve) {
                    window.Asc.scope.bibEntries = entries;

                    window.Asc.plugin.callCommand(function () {
                        var oDocument = Api.GetDocument();
                        var entries = Asc.scope.bibEntries;

                        var controls = oDocument.GetContentControlsByTag("ZOTERO_BIBLIOGRAPHY");
                        if (!controls || controls.length === 0) {
                            return;
                        }

                        var oBlockLvlSdt = controls[0];
                        oBlockLvlSdt.RemoveAllElements();

                        for (var i = 0; i < entries.length; i++) {
                            var oParagraph = Api.CreateParagraph();
                            oParagraph.SetFontFamily("Times New Roman");
                            oParagraph.SetIndLeft(720);        // 0.5" / 1.27 cm
                            oParagraph.SetIndFirstLine(-720);  // 悬挂缩进 (Hanging Indent)
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
                            oBlockLvlSdt.AddElement(oParagraph);
                        }
                    }, false, true, function () {
                        resolve(true);
                    });
                });
            });
        }).catch(function (err) {
            console.warn("Auto-upgrade error:", err);
        });
    }

    /**
     * Reformat all citations and bibliography across the entire document to a new CSL style.
     * Ubiquitously converts Narrative citations between any language/style representations
     * (e.g. 'Yan et al. (2026)' <-> '严等 (2026)' or 'Wang等 (2026)' or 'Schmidt u. a. (2024)') in place!
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
                    if (it.id) allKeysSet[it.id] = true;
                });
            });
            var allKeys = Object.keys(allKeysSet);
            if (!allKeys.length) return 0;

            return ZoteroAPI.getItemsByKeys(allKeys, 0, newStyle).then(function (items) {
                var itemMap = {};
                items.forEach(function (it) {
                    itemMap[it.key] = it;
                });

                var updatedFields = [];
                fields.forEach(function (f) {
                    var clusterItems = (f.citationItems || []).map(function (it) {
                        return itemMap[it.id] || { key: it.id, id: it.id, authors: "Unknown", year: "" };
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
                            var promises = updatedFields.map(function (f) {
                                return applyFontToField(f.FieldId);
                            });
                            Promise.all(promises).then(function () {
                                autoUpgradeBibliography(newStyle, allKeys).then(function () {
                                    resolve(updatedFields.length);
                                });
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
                    return f.Value && (f.Value.indexOf("ZOTERO_ITEM") !== -1 || f.Value.indexOf("ZOTERO_BIBL") !== -1);
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
        getCurrentField: getCurrentField,
        getAllCitationFields: getAllCitationFields,
        hasBibliographyControl: hasBibliographyControl,
        buildCitation: buildCitation,
        universalExtractNarrative: universalExtractNarrative,
        computeNarrativeAuthor: computeNarrativeAuthor,
        insertCitation: insertCitation,
        updateCitation: updateCitation,
        removeCitation: removeCitation,
        insertBibliography: insertBibliography,
        autoUpgradeBibliography: autoUpgradeBibliography,
        reformatDocumentStyle: reformatDocumentStyle,
        unlinkAll: unlinkAll
    };
})();
