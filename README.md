# 📊 Excel Clone 

A powerful, web-based spreadsheet application built using **HTML**, **CSS**, **JavaScript (jQuery)**, and **XLSX.js** to emulate Microsoft Excel-like functionalities in the browser.

> ✨ This project is not only visually close to MS Excel but also mimics many of its features like cell formatting, sheet management, save/export as `.json` and `.xlsx`, and more.

---

## 🚀 Features

### 🧮 Spreadsheet Functionalities
- Editable **grid of cells** with row and column headers.
- Click, type, select — just like real Excel.
- **Multiple sheet tabs** with support for:
  - Adding
  - Renaming
  - Deleting
- Customizable **cell styling**:
  - Bold, Italic, Underline
  - Font Size and Family
  - Text and Background Color
  - Left, Center, Right Alignment

### 📁 File Operations
- **Save** data as `.json` for persistent storage.
- **Export** as `.xlsx` using [`SheetJS`](https://sheetjs.com/).
- **Name-based cell selection** like `A1`, `B2`, etc.

### 🧭 Toolbar Tabs
- Navigation bar with tabbed menu system:
  - 🏠 Home (Clipboard, Font, Alignment)
  - ➕ Insert (Tables, Images — planned)
  - 🧩 Formulas (coming soon)
  - 📐 Layout and 📊 Data (placeholders for future)

### 🎨 Themes and UI
- Light/Dark **theme toggle** 🌙☀️.
- Scroll-synced column and row headers.
- Keyboard shortcuts: `Enter` in NameBox focuses specific cell.

### 🔁 Undo/Redo
- Stack-based undo and redo for content changes.

### 📤 Share Feature
- Dedicated `share.html` page to simulate sharing (UI).
- (Optional) Email integration logic via backend (planned).

---

## 🛠️ Technologies Used

- **HTML5 / CSS3**
- **JavaScript** (ES6)
- **jQuery**
- **SheetJS (XLSX.js)** for Excel file handling
- **Flexbox + Grid** layout


---

## 🧪 Project Structure

Excel-Clone-master/
├── index.html → Main spreadsheet interface
├── home.html → Home tab content
├── share.html → Share page UI
├── script.js → Main logic and interactivity
├── style.css → Stylesheet for layout and design
├── README.md → This file!
└── package.json → Metadata (optional usage)


## 🧰 How to Run

1. **Clone the Repository**  
   ```bash
   git clone https://github.com/Richa-Ranjan/excel_clone.git
   cd excel_clone/Excel-Clone-master
Open in Browser
Just open index.html in any modern browser (Chrome recommended).

To Save Excel File

Use the “Save As” button.

A .xlsx file will download to your system.

To Share

Click “Share” → Fill dummy mail field (real backend can be added).

💡 Unique Highlights:-
🧠 Dynamic sheet management with clean JSON structure.
🪄 Real-time formatting updates with a responsive toolbar.
📥 One-click download for Excel without a server.
🔄 Reversible edits with undo/redo just like real Office tools.
🌐 Simulated Share functionality with UI transition.
🎯 Cell selection from name box with scroll-to support.


✍️ Author
Richa Ranjan
💼 https://github.com/Richa-Ranjan


📃 License
This project is open-sourced for learning and demonstration purposes. Feel free to fork and extend!

🏁 Future Plans
✅ Formula engine (basic math like =A1+B1)
🔜 Charts and conditional formatting
🔜 Collaboration support
🔜 Backend file/email sharing (Node.js or Firebase)
