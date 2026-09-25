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
borrowbloc/
├── index.html       # Main single-file frontend & dashboard logic
├── README.md        # Project documentation
└── assets/          # (Optional) Mock screenshots or icons
