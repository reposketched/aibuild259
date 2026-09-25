# BorrowBloc

> *Neighborhood tool-sharing made frictionless. Stop buying duplicates; start sharing resources.*

BorrowBloc is a lightweight, zero-auth inventory and lending tracker designed to transform local neighborhoods into collaborative sharing economies. Built as a rapid hackathon project, it cuts down waste, saves money, and builds community trust by making local tool sharing as easy as clicking a button.

---

## Features

* **Instant Inventory Visibility:** Real-time dashboard showing available versus checked-out community tools (drills, ladders, wheelbarrows, and more).
* **Frictionless Checkout:** No clunky user accounts or passwords required. Simply click a tool, enter your name, and check it out instantly.
* **Google Sheets Backend:** Uses a connected Google Sheet as a live cloud database, allowing organizers to manage inventory or test live sync during demos effortlessly.
* **Instant Search & Filter:** Quickly filter through tools by name or category.
* **Responsive & Clean UI:** Styled with Tailwind CSS for a modern, approachable aesthetic.

---

## Tech Stack

* **Frontend:** HTML5, Vanilla JavaScript, Tailwind CSS (via CDN)
* **Icons:** Lucide Icons
* **Backend / Database:** Google Sheets API (via SheetBest / Google Apps Script wrapper)

---

## Project Structure

```text
aibuild259/
├── index.html          # Single-file frontend & dashboard logic (runs in demo mode out of the box)
├── apps-script/
│   └── Code.gs         # Google Apps Script backend for the Inventory sheet
├── specsheet.txt       # Hackathon spec
└── README.md           # Project documentation
```

---

## Quick Start

1. Open `index.html` in a browser. With no backend configured it runs on demo data.
2. To go live: create a Google Sheet with an `Inventory` tab (headers: ID, Tool Name, Category, Status, Borrower, Due Date).
3. Extensions → Apps Script → paste `apps-script/Code.gs` → Deploy as a Web app (Execute as: Me, Access: Anyone).
4. Paste the `/exec` URL into `SHEET_ENDPOINT` at the top of the script in `index.html`.
