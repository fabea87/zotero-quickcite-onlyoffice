/**
 * Zotero QuickCite - Application Controller
 * Manages purely keyboard-driven UI, narrative citations, colon-syntax page parsing ('author:42'),
 * streamlined 4 common Author-Date styles by default (APA, Chicago, MLA, GB/T 2005),
 * and a full Style Management UI to retain, reorder, and search styles.
 */
(function () {
    "use strict";

    // Four core common Author-Date styles enabled by default
    var DEFAULT_STARTER_STYLES = [
        { id: "apa", title: "APA Style 7th edition", format: "author-date" },
        { id: "chicago-author-date", title: "Chicago Manual of Style 18th (Author-Date)", format: "author-date" },
        { id: "modern-language-association", title: "MLA Handbook 9th edition", format: "author-date" },
        { id: "chinese-gb7714-2005-author-date", title: "GB/T 7714-2005 (Author-Date)", format: "author-date" }
    ];

    // All known genuine styles available in Zotero Desktop
    var ALL_ZOTERO_STYLES = [
        { id: "apa", title: "APA Style 7th edition", format: "author-date" },
        { id: "chicago-author-date", title: "Chicago Manual of Style 18th (Author-Date)", format: "author-date" },
        { id: "modern-language-association", title: "MLA Handbook 9th edition", format: "author-date" },
        { id: "chinese-gb7714-2005-author-date", title: "GB/T 7714-2005 (Author-Date)", format: "author-date" },
        { id: "system", title: "System (Elsevier Applied Linguistics)", format: "author-date" },
        { id: "elsevier-harvard", title: "Elsevier (Author-Date / Harvard)", format: "author-date" },
        { id: "harvard-cite-them-right", title: "Cite Them Right 12th (Harvard)", format: "author-date" },
        { id: "american-sociological-association", title: "ASA Style Guide", format: "author-date" },
        { id: "american-political-science-association", title: "APSA Style Manual", format: "author-date" },
        { id: "chicago-notes-bibliography", title: "Chicago 18th (Notes & Bibliography)", format: "note" },
        { id: "chicago-shortened-notes-bibliography", title: "Chicago 18th (Shortened Notes)", format: "note" },
        { id: "mhra-notes", title: "MHRA Style Guide 4th (Notes)", format: "note" },
        { id: "ieee", title: "IEEE Reference Guide", format: "numeric" },
        { id: "nature", title: "Nature", format: "numeric" },
        { id: "nlm-citation-sequence", title: "NLM / Vancouver (Citation-Sequence)", format: "numeric" },
        { id: "american-chemical-society", title: "ACS Guide", format: "numeric" },
        { id: "american-medical-association", title: "AMA Manual of Style 11th", format: "numeric" },
        { id: "acta-paediatrica", title: "Acta Paediatrica", format: "numeric" }
    ];

    var state = {
        mode: "insert",
        editingField: null,
        selectedItems: [],
        searchResults: [],
        selectedIndex: -1,
        activeStyle: localStorage.getItem("zotero_active_style") || "apa",
        retainedStyles: [],
        activeLibrary: "0",
        formatMode: "parenthetical",
        searchTimer: null
    };

    var els = {};

    function getElements() {
        els.statusDot = document.getElementById("statusDot");
        els.statusText = document.getElementById("statusText");
        els.styleSelect = document.getElementById("styleSelect");
        els.librarySelect = document.getElementById("librarySelect");
        els.modeBanner = document.getElementById("modeBanner");
        els.modeTitle = document.getElementById("modeTitle");
        els.selectedChips = document.getElementById("selectedChips");
        els.searchInput = document.getElementById("searchInput");
        els.searchClear = document.getElementById("searchClear");
        els.resultsList = document.getElementById("resultsList");
        els.inputLocator = document.getElementById("inputLocator");
        els.inputPrefix = document.getElementById("inputPrefix");
        els.inputSuffix = document.getElementById("inputSuffix");
        els.segmentBtns = document.querySelectorAll(".segment-btn");
        els.livePreview = document.getElementById("livePreview");
        els.btnPrimary = document.getElementById("btnPrimary");
        els.btnFootnote = document.getElementById("btnFootnote");
        els.btnBibliography = document.getElementById("btnBibliography");
        els.btnUnlink = document.getElementById("btnUnlink");
        els.btnDeleteCitation = document.getElementById("btnDeleteCitation");
        els.btnClose = document.getElementById("btnClose");
        els.toast = document.getElementById("toast");

        // Style Manager Elements
        els.styleModal = document.getElementById("styleModal");
        els.retainedStylesList = document.getElementById("retainedStylesList");
        els.styleSearchInput = document.getElementById("styleSearchInput");
        els.styleResultsList = document.getElementById("styleResultsList");
        els.btnCloseStyleModal = document.getElementById("btnCloseStyleModal");
    }

    function showToast(msg, isError) {
        if (!els.toast) return;
        els.toast.textContent = msg;
        els.toast.className = "toast show" + (isError ? " toast-error" : "");
        setTimeout(function () {
            els.toast.className = "toast";
        }, 2800);
    }

    function closeModal() {
        try {
            window.Asc.plugin.executeCommand("close", "");
        } catch (e) {
            console.warn("Failed to close modal:", e);
        }
    }

    // ========== Initialization ==========

    window.Asc.plugin.init = function () {
        getElements();
        loadRetainedStyles();
        setupEvents();
        setupContextMenu();

        initStyleDropdown();

        ZoteroAPI.ping().then(function (connected) {
            updateConnectionStatus(connected);
            if (connected) {
                loadGroups();
            }
        });

        CitationManager.getCurrentField().then(function (field) {
            if (field) {
                setupEditMode(field);
            } else {
                setupInsertMode();
            }
        }).catch(function (err) {
            console.warn("Error detecting current field:", err);
            setupInsertMode();
        });

        // Detect if bibliography already exists and rename button to "Update Bibliography"
        CitationManager.hasBibliographyControl().then(function (hasBib) {
            updateBibButtonState(hasBib);
        });
    };

    window.Asc.plugin.button = function () {
        closeModal();
    };

    function setupContextMenu() {
        if (window.Asc && window.Asc.ButtonContextMenu && window.Asc.Buttons) {
            try {
                var editBtn = new window.Asc.ButtonContextMenu();
                editBtn.text = "✏️ Edit Zotero Citation";
                editBtn.addCheckers("Target", "Selection");
                editBtn.attachOnClick(function () {
                    CitationManager.getCurrentField().then(function (field) {
                        if (field) setupEditMode(field);
                    });
                });
                window.Asc.Buttons.registerContextMenu();
            } catch (e) {
                console.warn("Context menu registration:", e);
            }
        }
    }

    function updateConnectionStatus(connected) {
        if (connected) {
            els.statusDot.className = "status-dot connected";
            els.statusText.textContent = "Zotero Connected";
        } else {
            els.statusDot.className = "status-dot disconnected";
            els.statusText.innerHTML = "Zotero Offline (<a href='#' id='btnRetry' style='color:inherit;text-decoration:underline;'>Retry</a>)";
            var btnRetry = document.getElementById("btnRetry");
            if (btnRetry) {
                btnRetry.onclick = function (e) {
                    e.preventDefault();
                    els.statusText.textContent = "Connecting...";
                    ZoteroAPI.ping().then(updateConnectionStatus);
                };
            }
        }
    }

    function loadGroups() {
        ZoteroAPI.getGroups().then(function (groups) {
            if (Array.isArray(groups) && groups.length > 0) {
                els.librarySelect.style.display = "inline-block";
                groups.forEach(function (g) {
                    var opt = document.createElement("option");
                    opt.value = String(g.id);
                    opt.textContent = g.name;
                    els.librarySelect.appendChild(opt);
                });
            }
        });
    }

    // ========== Retained Styles & Style Management UI ==========

    function loadRetainedStyles() {
        try {
            var saved = JSON.parse(localStorage.getItem("zotero_retained_styles") || "null");
            if (Array.isArray(saved) && saved.length > 0) {
                state.retainedStyles = saved;
                return;
            }
        } catch (e) {}

        // Default to the 4 common Author-Date styles
        state.retainedStyles = DEFAULT_STARTER_STYLES.slice();
    }

    function saveRetainedStyles() {
        localStorage.setItem("zotero_retained_styles", JSON.stringify(state.retainedStyles));
    }

    function initStyleDropdown() {
        els.styleSelect.innerHTML = "";

        state.retainedStyles.forEach(function (s) {
            var opt = document.createElement("option");
            opt.value = s.id;
            opt.textContent = s.title;
            if (s.id === state.activeStyle) opt.selected = true;
            els.styleSelect.appendChild(opt);
        });

        // Add action item to open the Style Manager
        var manageOpt = document.createElement("option");
        manageOpt.value = "__manage_styles__";
        manageOpt.textContent = "⚙️ Manage Styles / Add More...";
        els.styleSelect.appendChild(manageOpt);
    }

    function switchDocumentStyle(newStyle, styleTitle) {
        state.activeStyle = newStyle;
        localStorage.setItem("zotero_active_style", newStyle);

        showToast("Reformatting document to " + (styleTitle || newStyle) + " in Times New Roman...");

        CitationManager.reformatDocumentStyle(newStyle).then(function (count) {
            if (count > 0) {
                showToast("Updated " + count + " citation(s) and bibliography to " + (styleTitle || newStyle) + "!");
            } else {
                showToast("Style set to " + (styleTitle || newStyle));
            }
            updateLivePreview();
            performSearch(cleanSearchQuery(els.searchInput.value));
        }).catch(function (err) {
            showToast("Reformat error: " + err.message, true);
        });
    }

    function openStyleModal() {
        els.styleModal.style.display = "flex";
        els.styleSearchInput.value = "";
        renderRetainedStylesManager();
        renderSearchableOtherStyles("");
        els.styleSearchInput.focus();
    }

    function closeStyleModal() {
        els.styleModal.style.display = "none";
        initStyleDropdown();
        els.searchInput.focus();
    }

    function moveRetainedStyle(index, delta) {
        var newIndex = index + delta;
        if (newIndex < 0 || newIndex >= state.retainedStyles.length) return;

        var item = state.retainedStyles.splice(index, 1)[0];
        state.retainedStyles.splice(newIndex, 0, item);

        saveRetainedStyles();
        renderRetainedStylesManager();
        initStyleDropdown();
    }

    function removeRetainedStyle(index) {
        if (state.retainedStyles.length <= 1) {
            showToast("You must keep at least one style in the dropdown.", true);
            return;
        }

        var removed = state.retainedStyles.splice(index, 1)[0];
        saveRetainedStyles();
        renderRetainedStylesManager();
        renderSearchableOtherStyles(els.styleSearchInput.value);
        initStyleDropdown();

        // If removed style was the active one, switch to the first retained style
        if (removed.id === state.activeStyle && state.retainedStyles.length > 0) {
            switchDocumentStyle(state.retainedStyles[0].id, state.retainedStyles[0].title);
        }

        showToast("Removed '" + removed.title + "' from dropdown.");
    }

    function addStyleToRetained(styleObj) {
        var exists = state.retainedStyles.some(function (s) { return s.id === styleObj.id; });
        if (exists) {
            showToast("Style is already in your dropdown list.");
            return;
        }

        state.retainedStyles.push(styleObj);
        saveRetainedStyles();
        renderRetainedStylesManager();
        renderSearchableOtherStyles(els.styleSearchInput.value);
        initStyleDropdown();

        // Immediately switch document to the newly added style
        switchDocumentStyle(styleObj.id, styleObj.title);
        showToast("Added '" + styleObj.title + "' and applied to document!");
    }

    function renderRetainedStylesManager() {
        els.retainedStylesList.innerHTML = "";

        state.retainedStyles.forEach(function (s, idx) {
            var row = document.createElement("div");
            row.className = "retained-style-row" + (s.id === state.activeStyle ? " active" : "");

            var info = document.createElement("div");
            info.className = "retained-style-info";

            var title = document.createElement("span");
            title.className = "retained-style-title";
            title.textContent = s.title;

            var idBadge = document.createElement("span");
            idBadge.className = "style-item-id";
            idBadge.textContent = "(" + s.id + ")";

            info.appendChild(title);
            info.appendChild(idBadge);

            if (s.id === state.activeStyle) {
                var activeBadge = document.createElement("span");
                activeBadge.className = "style-item-badge";
                activeBadge.textContent = "● Active in Document";
                info.appendChild(activeBadge);
            }

            var actions = document.createElement("div");
            actions.className = "retained-style-actions";

            // Up button
            var btnUp = document.createElement("button");
            btnUp.className = "btn-icon";
            btnUp.innerHTML = "↑";
            btnUp.title = "Move Up";
            btnUp.disabled = (idx === 0);
            btnUp.onclick = function (e) {
                e.stopPropagation();
                moveRetainedStyle(idx, -1);
            };

            // Down button
            var btnDown = document.createElement("button");
            btnDown.className = "btn-icon";
            btnDown.innerHTML = "↓";
            btnDown.title = "Move Down";
            btnDown.disabled = (idx === state.retainedStyles.length - 1);
            btnDown.onclick = function (e) {
                e.stopPropagation();
                moveRetainedStyle(idx, 1);
            };

            // Set Active button (if not already active)
            if (s.id !== state.activeStyle) {
                var btnUse = document.createElement("button");
                btnUse.className = "btn btn-default btn-sm";
                btnUse.textContent = "Use";
                btnUse.title = "Apply this style to document";
                btnUse.onclick = function (e) {
                    e.stopPropagation();
                    switchDocumentStyle(s.id, s.title);
                    renderRetainedStylesManager();
                };
                actions.appendChild(btnUse);
            }

            // Remove button
            var btnDel = document.createElement("button");
            btnDel.className = "btn-icon";
            btnDel.innerHTML = "✕";
            btnDel.title = "Remove from dropdown";
            btnDel.onclick = function (e) {
                e.stopPropagation();
                removeRetainedStyle(idx);
            };

            actions.appendChild(btnUp);
            actions.appendChild(btnDown);
            actions.appendChild(btnDel);

            row.appendChild(info);
            row.appendChild(actions);

            els.retainedStylesList.appendChild(row);
        });
    }

    function renderSearchableOtherStyles(query) {
        query = (query || "").toLowerCase().trim();
        var retainedIds = state.retainedStyles.map(function (s) { return s.id; });

        // Available other styles not currently retained
        var pool = ALL_ZOTERO_STYLES.filter(function (s) {
            return retainedIds.indexOf(s.id) === -1;
        });

        var matches = pool;
        if (query) {
            matches = pool.filter(function (s) {
                return s.title.toLowerCase().indexOf(query) !== -1 ||
                       s.id.toLowerCase().indexOf(query) !== -1;
            });
        }

        els.styleResultsList.innerHTML = "";
        if (!matches.length) {
            els.styleResultsList.innerHTML = "<div class='empty-message' style='padding:16px;'>No other matching styles found. All available styles are in your list.</div>";
            return;
        }

        matches.forEach(function (s) {
            var item = document.createElement("div");
            item.className = "style-item";

            var left = document.createElement("div");
            left.className = "style-item-header";

            var title = document.createElement("span");
            title.className = "style-item-title";
            title.textContent = s.title;

            var idText = document.createElement("span");
            idText.className = "style-item-id";
            idText.textContent = "(" + s.id + ")";

            left.appendChild(title);
            left.appendChild(idText);

            var addBtn = document.createElement("button");
            addBtn.className = "btn btn-primary btn-sm";
            addBtn.textContent = "+ Add to Dropdown";
            addBtn.onclick = function (e) {
                e.stopPropagation();
                addStyleToRetained(s);
            };

            item.appendChild(left);
            item.appendChild(addBtn);

            els.styleResultsList.appendChild(item);
        });
    }

    // ========== Citation Mode Setup ==========

    function setupInsertMode() {
        state.mode = "insert";
        state.editingField = null;
        state.selectedItems = [];

        els.modeBanner.style.display = "none";
        els.btnDeleteCitation.style.display = "none";
        els.btnPrimary.textContent = "Insert Citation (↵)";
        renderChips();

        els.searchInput.focus();
        performSearch("");
    }

    function setupEditMode(field) {
        state.mode = "edit";
        state.editingField = field;

        els.modeBanner.style.display = "flex";
        els.modeTitle.textContent = "Editing In-Text Citation";
        els.btnDeleteCitation.style.display = "inline-flex";
        els.btnPrimary.textContent = "Update Citation (↵)";

        var firstItem = (field.citationItems && field.citationItems[0]) || {};

        if (firstItem.locator) els.inputLocator.value = firstItem.locator;
        if (firstItem.prefix) els.inputPrefix.value = firstItem.prefix;
        if (firstItem.suffix) els.inputSuffix.value = firstItem.suffix;

        if (firstItem["suppress-author"]) {
            setFormatMode("suppress");
        } else {
            setFormatMode("parenthetical");
        }

        var keys = [];
        (field.citationItems || []).forEach(function (citItem) {
            if (citItem.id) keys.push(citItem.id);
        });

        if (keys.length > 0) {
            ZoteroAPI.getItemsByKeys(keys, 0, state.activeStyle).then(function (items) {
                state.selectedItems = items;
                renderChips();
                updateLivePreview();
            });
        }

        els.searchInput.placeholder = "Type to add another reference to this cluster...";
        els.searchInput.focus();
    }

    function setFormatMode(mode) {
        state.formatMode = mode;
        els.segmentBtns.forEach(function (btn) {
            if (btn.getAttribute("data-mode") === mode) {
                btn.classList.add("active");
            } else {
                btn.classList.remove("active");
            }
        });
        updateLivePreview();
    }

    function updateLivePreview() {
        var item = state.selectedItems[0] || (state.selectedIndex >= 0 ? state.searchResults[state.selectedIndex] : null);
        if (!item) {
            els.livePreview.innerHTML = "<span class='preview-hint'>Preview will appear when an item is selected</span>";
            return;
        }

        var options = {
            locator: els.inputLocator.value.trim(),
            prefix: els.inputPrefix.value,
            suffix: els.inputSuffix.value,
            formatMode: state.formatMode,
            style: state.activeStyle
        };

        var built = CitationManager.buildCitation(state.selectedItems.length ? state.selectedItems : [item], options);

        if (built.isNarrative && built.narrativeAuthor) {
            els.livePreview.innerHTML = "<b>Preview:</b> <span class='preview-author'>" + built.narrativeAuthor + "</span> " + "<span class='preview-field'>" + built.formattedText + "</span>";
        } else {
            els.livePreview.innerHTML = "<b>Preview:</b> <span class='preview-field'>" + built.formattedText + "</span>";
        }
    }

    // ========== Chips (Multi-citation cluster) ==========

    function renderChips() {
        els.selectedChips.innerHTML = "";
        if (!state.selectedItems.length) {
            els.selectedChips.style.display = "none";
            updateLivePreview();
            return;
        }

        els.selectedChips.style.display = "flex";
        state.selectedItems.forEach(function (item, idx) {
            var chip = document.createElement("div");
            chip.className = "chip";

            var label = document.createElement("span");
            var author = item.authors || "Item";
            var yr = item.year ? (" (" + item.year + ")") : "";
            label.textContent = author + yr;

            var removeBtn = document.createElement("button");
            removeBtn.className = "chip-remove";
            removeBtn.innerHTML = "×";
            removeBtn.onclick = function (e) {
                e.stopPropagation();
                state.selectedItems.splice(idx, 1);
                renderChips();
                updateLivePreview();
            };

            chip.appendChild(label);
            chip.appendChild(removeBtn);
            els.selectedChips.appendChild(chip);
        });

        updateLivePreview();
    }

    // ========== Searching & Colon Page Parsing ==========

    function cleanSearchQuery(raw) {
        if (!raw) return "";
        var colonMatch = raw.match(/:([0-9\-]+)$/);
        var pageMatch = raw.match(/\s+p\.?\s*([0-9\-]+)$/i);

        if (colonMatch) {
            els.inputLocator.value = colonMatch[1];
            updateLivePreview();
            return raw.substring(0, colonMatch.index).trim();
        } else if (pageMatch) {
            els.inputLocator.value = pageMatch[1];
            updateLivePreview();
            return raw.substring(0, pageMatch.index).trim();
        }
        return raw.trim();
    }

    var activeSearchSeq = 0;

    function onSearchInput(e) {
        var rawVal = e.target.value;
        if (els.searchClear) {
            els.searchClear.style.display = rawVal.length ? "block" : "none";
        }

        var query = cleanSearchQuery(rawVal);

        if (state.searchTimer) clearTimeout(state.searchTimer);
        state.searchTimer = setTimeout(function () {
            performSearch(query);
        }, 260);
    }

    function performSearch(query) {
        var thisSeq = ++activeSearchSeq;
        els.resultsList.innerHTML = "<div class='loading-spinner'>Searching Zotero...</div>";

        ZoteroAPI.searchItems(query, state.activeLibrary, state.activeStyle).then(function (items) {
            if (thisSeq !== activeSearchSeq) return; // Discard superseded response!

            state.searchResults = items || [];
            state.selectedIndex = state.searchResults.length > 0 ? 0 : -1;
            renderResults();
            updateLivePreview();
        }).catch(function () {
            if (thisSeq !== activeSearchSeq) return;

            ZoteroAPI.ping().then(function (isAlive) {
                if (isAlive) {
                    els.resultsList.innerHTML = "<div class='empty-message'>No matching references found. Try adjusting keywords.</div>";
                } else {
                    els.resultsList.innerHTML = "<div class='empty-message'>Unable to reach Zotero Desktop. Please ensure Zotero is running.</div>";
                }
            });
        });
    }

    function renderResults() {
        els.resultsList.innerHTML = "";
        if (!state.searchResults.length) {
            els.resultsList.innerHTML = "<div class='empty-message'>No matching references found.</div>";
            return;
        }

        state.searchResults.forEach(function (item, idx) {
            var row = document.createElement("div");
            row.className = "result-item" + (idx === state.selectedIndex ? " selected" : "");

            var header = document.createElement("div");
            header.className = "result-header";

            var authorYear = document.createElement("span");
            authorYear.className = "result-author-year";
            authorYear.textContent = (item.authors || "Unknown") + (item.year ? (" (" + item.year + ")") : "");

            var badge = document.createElement("span");
            badge.className = "result-badge";
            badge.textContent = item.itemType || "item";

            header.appendChild(authorYear);
            header.appendChild(badge);

            var title = document.createElement("div");
            title.className = "result-title";
            title.textContent = item.title;

            var meta = document.createElement("div");
            meta.className = "result-meta";
            meta.textContent = item.publication ? (item.publication + (item.citation ? " • " + item.citation : "")) : (item.citation || "");

            row.appendChild(header);
            row.appendChild(title);
            if (item.publication || item.citation) row.appendChild(meta);

            row.onclick = function () {
                state.selectedIndex = idx;
                highlightSelectedRow();
                selectItem(item);
            };

            row.ondblclick = function () {
                state.selectedIndex = idx;
                selectItem(item);
                submitCitation(false);
            };

            els.resultsList.appendChild(row);
        });

        ensureVisible(state.selectedIndex);
    }

    function highlightSelectedRow() {
        var rows = els.resultsList.querySelectorAll(".result-item");
        rows.forEach(function (r, idx) {
            if (idx === state.selectedIndex) {
                r.classList.add("selected");
            } else {
                r.classList.remove("selected");
            }
        });
        updateLivePreview();
    }

    function ensureVisible(idx) {
        var rows = els.resultsList.querySelectorAll(".result-item");
        if (rows[idx]) {
            rows[idx].scrollIntoView({ block: "nearest" });
        }
    }

    function selectItem(item) {
        var existingIdx = -1;
        for (var i = 0; i < state.selectedItems.length; i++) {
            if (state.selectedItems[i].key === item.key) {
                existingIdx = i;
                break;
            }
        }
        if (existingIdx === -1) {
            state.selectedItems.push(item);
            renderChips();
        }
    }

    // ========== Citation Submission ==========

    function submitCitation(isFootnote) {
        if (!state.selectedItems.length && state.selectedIndex >= 0 && state.searchResults[state.selectedIndex]) {
            state.selectedItems.push(state.searchResults[state.selectedIndex]);
        }

        if (!state.selectedItems.length) {
            showToast("Please select a reference to insert.", true);
            return;
        }

        var options = {
            locator: els.inputLocator.value.trim(),
            prefix: els.inputPrefix.value,
            suffix: els.inputSuffix.value,
            formatMode: state.formatMode,
            style: state.activeStyle,
            isFootnote: isFootnote
        };

        var citationData = CitationManager.buildCitation(state.selectedItems, options);

        if (state.mode === "edit" && state.editingField) {
            CitationManager.updateCitation(state.editingField.fieldId, citationData, state.activeStyle).then(function () {
                closeModal();
            }).catch(function (err) {
                showToast("Update failed: " + err.message, true);
            });
        } else {
            CitationManager.insertCitation(citationData, isFootnote, state.activeStyle).then(function () {
                closeModal();
            }).catch(function (err) {
                showToast("Insert failed: " + err.message, true);
            });
        }
    }

    function deleteCurrentCitation() {
        if (state.editingField) {
            CitationManager.removeCitation(state.editingField.fieldId, state.activeStyle).then(function () {
                closeModal();
            });
        }
    }

    function updateBibButtonState(hasBib) {
        if (!els.btnBibliography) return;
        if (hasBib) {
            els.btnBibliography.textContent = "Update Bibliography";
            els.btnBibliography.title = "Update existing reference list in document (Alt+B)";
        } else {
            els.btnBibliography.textContent = "Insert Bibliography";
            els.btnBibliography.title = "Insert reference list in Times New Roman with hanging indents (Alt+B)";
        }
    }

    function handleInsertBibliography() {
        var isUpdate = (els.btnBibliography.textContent.indexOf("Update") !== -1);
        els.btnBibliography.disabled = true;
        els.btnBibliography.textContent = isUpdate ? "Updating..." : "Formatting...";

        CitationManager.insertBibliography(state.activeStyle).then(function (count) {
            updateBibButtonState(true);
            showToast(isUpdate ? "Bibliography updated (" + count + " entries)!" : "Bibliography inserted in Times New Roman with hanging indents (" + count + " entries)!");
            setTimeout(closeModal, 600);
        }).catch(function (err) {
            els.btnBibliography.disabled = false;
            updateBibButtonState(isUpdate);
            showToast("Bibliography error: " + err.message, true);
        });
    }

    function handleUnlink() {
        if (confirm("Unlink all Zotero citations in this document? This converts all fields into plain text.")) {
            CitationManager.unlinkAll().then(function (count) {
                showToast("Unlinked " + count + " citation(s).");
                setTimeout(closeModal, 600);
            });
        }
    }

    // ========== Comprehensive Keyboard Shortcuts ==========

    function setupEvents() {
        els.searchInput.addEventListener("input", onSearchInput);
        els.inputLocator.addEventListener("input", updateLivePreview);
        els.inputPrefix.addEventListener("input", updateLivePreview);
        els.inputSuffix.addEventListener("input", updateLivePreview);

        els.segmentBtns.forEach(function (btn) {
            btn.addEventListener("click", function () {
                var mode = btn.getAttribute("data-mode");
                setFormatMode(mode);
            });
        });

        if (els.searchClear) {
            els.searchClear.onclick = function () {
                els.searchInput.value = "";
                els.searchClear.style.display = "none";
                performSearch("");
                els.searchInput.focus();
            };
        }

        // Global pure-keyboard handling
        document.addEventListener("keydown", function (e) {
            var key = e.key;

            if (key === "Escape") {
                e.preventDefault();
                if (els.styleModal.style.display === "flex") {
                    closeStyleModal();
                } else {
                    closeModal();
                }
                return;
            }

            if (els.styleModal.style.display === "flex") {
                return;
            }

            if (e.altKey) {
                var lowerKey = key.toLowerCase();
                if (lowerKey === "n") {
                    e.preventDefault();
                    setFormatMode("narrative");
                    return;
                } else if (lowerKey === "p") {
                    e.preventDefault();
                    setFormatMode("parenthetical");
                    return;
                } else if (lowerKey === "y") {
                    e.preventDefault();
                    setFormatMode("suppress");
                    return;
                } else if (lowerKey === "l") {
                    e.preventDefault();
                    els.inputLocator.focus();
                    els.inputLocator.select();
                    return;
                } else if (lowerKey === "s") {
                    e.preventDefault();
                    els.searchInput.focus();
                    els.searchInput.select();
                    return;
                } else if (lowerKey === "b") {
                    e.preventDefault();
                    handleInsertBibliography();
                    return;
                }
            }

            if ((e.metaKey || e.ctrlKey) && (key === "Backspace" || key === "Delete")) {
                if (state.mode === "edit" && state.editingField) {
                    e.preventDefault();
                    deleteCurrentCitation();
                    return;
                }
            }

            if (key === "Tab") {
                var cycle = [els.searchInput, els.inputLocator, els.inputPrefix, els.inputSuffix];
                var curIdx = cycle.indexOf(document.activeElement);

                if (curIdx !== -1) {
                    e.preventDefault();
                    var nextIdx;
                    if (e.shiftKey) {
                        nextIdx = (curIdx - 1 + cycle.length) % cycle.length;
                    } else {
                        nextIdx = (curIdx + 1) % cycle.length;
                    }
                    cycle[nextIdx].focus();
                    if (typeof cycle[nextIdx].select === "function") {
                        cycle[nextIdx].select();
                    }
                    return;
                }
            }

            if (key === "ArrowDown") {
                if (document.activeElement === els.searchInput || document.activeElement === document.body) {
                    e.preventDefault();
                    if (state.searchResults.length > 0) {
                        state.selectedIndex = (state.selectedIndex + 1) % state.searchResults.length;
                        highlightSelectedRow();
                        ensureVisible(state.selectedIndex);
                    }
                    return;
                }
            }

            if (key === "ArrowUp") {
                if (document.activeElement === els.searchInput || document.activeElement === document.body) {
                    e.preventDefault();
                    if (state.searchResults.length > 0) {
                        state.selectedIndex = (state.selectedIndex - 1 + state.searchResults.length) % state.searchResults.length;
                        highlightSelectedRow();
                        ensureVisible(state.selectedIndex);
                    }
                    return;
                }
            }

            if (key === "Enter") {
                if (document.activeElement === els.btnBibliography || document.activeElement === els.btnUnlink) {
                    return;
                }
                e.preventDefault();
                submitCitation(e.shiftKey);
                return;
            }
        });

        els.btnPrimary.onclick = function () { submitCitation(false); };
        els.btnFootnote.onclick = function () { submitCitation(true); };
        els.btnDeleteCitation.onclick = deleteCurrentCitation;
        els.btnBibliography.onclick = handleInsertBibliography;
        els.btnUnlink.onclick = handleUnlink;
        els.btnClose.onclick = closeModal;

        els.styleSelect.onchange = function (e) {
            var val = e.target.value;
            if (val === "__manage_styles__") {
                openStyleModal();
            } else {
                var selectedOpt = els.styleSelect.options[els.styleSelect.selectedIndex];
                switchDocumentStyle(val, selectedOpt ? selectedOpt.textContent : val);
            }
        };

        els.librarySelect.onchange = function (e) {
            state.activeLibrary = e.target.value;
            performSearch(cleanSearchQuery(els.searchInput.value));
        };

        els.styleSearchInput.addEventListener("input", function (e) {
            renderSearchableOtherStyles(e.target.value);
        });

        els.btnCloseStyleModal.onclick = closeStyleModal;
    }
})();
