# Modern File Search Dashboard

A **browser-based file search and report tool** built with **HTML, Tailwind CSS, JavaScript, SweetAlert2, pdf.js, and SheetJS**.  
Designed for **quick keyword search in files and PDFs**, with optional **matched file copying**, reporting, and a modern dashboard interface. Fully **Vercel-ready**.

---

## Features

- **Excel template-based search:** Upload `.xlsx` files with search keywords (header row is ignored).
- **Folder selection:** Pick **source** and **destination** folders using the modern File System Access API.
- **PDF support:** Search the **first page of PDFs** for keywords.
- **Duplicate detection:** Automatically identifies files with duplicate content.
- **Copy matched files:** Optionally copy all files that match search keywords to a destination folder.
- **Report generation:** Export results as `.xlsx` report including keywords, results, and file paths.
- **Real-time progress & logging:** Live log panel with auto-scroll, progress bar, and summary stats.
- **Modern UI:** Tailwind 3, card-style search results, responsive 2-column layout, and dark mode friendly.
- **Drag & drop template upload** with fallback file input.
- **Vercel-ready:** Works on modern browsers without a backend.

---

## Getting Started

### 1. Clone or Download

```bash
git clone https://github.com/yourusername/modern-file-search-dashboard.git
cd modern-file-search-dashboard
