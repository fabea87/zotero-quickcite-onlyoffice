# Zotero QuickCite for ONLYOFFICE

A fast, Word-style **floating search & insert popup** for ONLYOFFICE Desktop Editors with **bidirectional communication** to local Zotero Desktop (`http://127.0.0.1:23119`).

---

## 🚀 Easy Installation (Share with Colleagues)

### Option A: 1-Click GUI Install (Recommended for Everyone)
*No terminal or command line needed. Works identically on Windows, macOS, and Linux.*

1. Download or send the **`ZoteroQuickCite.plugin`** file to your colleague.
2. In **ONLYOFFICE Desktop Editors**, open any document (or create a new one).
3. On the top ribbon, click the **Plugins** tab &rarr; click **Plugin Manager** (the marketplace icon).
4. Click **"Install plugin manually"** (or **"Add plugin"** in settings).
5. Select the **`ZoteroQuickCite.plugin`** file.
6. Done! The red **Zotero QuickCite** icon will appear on the ribbon.

---

### Option B: Automatic Script Install (macOS / Windows / Linux)
If sharing via Git or ZIP archive:

```bash
# macOS / Linux
python3 install.py

# Windows (Command Prompt or PowerShell)
python install.py
```

The script automatically detects your OS and exact ONLYOFFICE version (Native, Flatpak, Snap, AppData), installs into the proper directory, and configures the plugin.

---

## ⌨️ Purely Keyboard-Driven Workflow

You never need to touch the mouse once the dialog opens:

| Action | Shortcut / Syntax |
| :--- | :--- |
| **Instant Focus** | Opens with cursor automatically in search. |
| **Search & Page Syntax** | Type `author:page` (e.g. `smith:42` or `panadero:15-18`) &rarr; auto-extracts page! |
| **Navigate Results** | `↓` (Down Arrow) / `↑` (Up Arrow) |
| **Cycle Focus** | `Tab` / `Shift + Tab` cycles `Search` &rarr; `Page` &rarr; `Prefix` &rarr; `Suffix` with text auto-selection. |
| **Parenthetical Mode** | `Alt + P` &rarr; `(Author, Year)` |
| **Narrative Mode** | `Alt + N` &rarr; `Author (Year)` (writes author in running text, e.g. `Yan et al. (2026)`). |
| **Year Only Mode** | `Alt + Y` &rarr; `(Year only)` (traditional suppress author). |
| **Jump to Page** | `Alt + L` &rarr; focus Page/Locator input directly. |
| **Jump to Search** | `Alt + S` &rarr; focus Search input directly. |
| **Insert Citation** | `Enter` &rarr; inserts and **closes modal immediately**. |
| **Footnote Citation** | `Shift + Enter` &rarr; creates footnote and inserts citation. |
| **Insert / Update Bibliography** | `Alt + B` &rarr; inserts or updates reference list. |
| **Delete Citation** | `Cmd + Backspace` (in Edit Mode) &rarr; removes citation from document. |
| **Cancel** | `Esc` &rarr; dismisses popup. |

---

## 📚 Key Features

* **100% Regulated by Zotero Output**: In-text citations and bibliographies match Zotero's exact CSL output without manual string manipulation or glued words (`&` vs `and`, `et al.`, spacing).
* **Strict Times New Roman**: Every in-text citation, narrative prefix, footnote, and bibliography paragraph is formatted in **Times New Roman**.
* **Automatic Bibliography Upgrades**:
  * The button displays **"Insert Bibliography"** on first use.
  * Once a bibliography exists, the button dynamically renames to **"Update Bibliography"**.
  * Adding, editing, or deleting any in-text citation automatically updates the existing bibliography in place!
* **Hanging Indent (悬挂缩进)**:
  Every bibliography entry is formatted with standard `SetIndLeft(720)` and `SetIndFirstLine(-720)` (0.5" / 1.27 cm hanging indent), with italics preserved on journal titles, book titles, and volume numbers.
* **Streamlined Author-Date Styles**:
  Starts by default with the 4 most common author-date styles:
  * **APA Style 7th edition** (`apa`)
  * **Chicago Manual of Style 18th (Author-Date)** (`chicago-author-date`)
  * **MLA Handbook 9th edition** (`modern-language-association`)
  * **GB/T 7714-2005 (Author-Date)** (`chinese-gb7714-2005-author-date`)
* **Style Management UI**:
  Click **`⚙️ Manage Styles / Add More...`** to reorder your quick dropdown with `↑` and `↓`, remove unused styles, or search and add any style from Zotero Desktop with real-time validation.
* **Right-Click Context Menu**:
  Right-click any citation in the document &rarr; select **`✏️ Edit Zotero Citation`** to open in Edit Mode.
* **Word (`.docx`) Compatibility**:
  Uses Microsoft Word's standard `ADDIN ZOTERO_ITEM CSL_CITATION` fields. Documents created in ONLYOFFICE can be edited seamlessly in Microsoft Word or LibreOffice.

---

## ⚙️ Requirements

* **ONLYOFFICE Desktop Editors** v7.3.3 or higher.
* **Zotero Desktop** (Zotero 6 or Zotero 7) running locally.
