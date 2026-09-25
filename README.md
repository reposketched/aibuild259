# BorrowBloc

> *Neighborhood tool-sharing made frictionless. Stop buying duplicates; start sharing resources.*

BorrowBloc is a lightweight, zero-auth inventory and lending tracker designed to transform local neighborhoods into collaborative sharing economies. Built as a rapid hackathon project, it cuts down waste, saves money, and builds community trust by making local tool sharing as easy as clicking a button.

---

## Features

* **Instant Inventory Visibility:** Real-time dashboard showing available versus checked-out community tools (drills, ladders, wheelbarrows, and more).
* **Frictionless Checkout:** No passwords. Pick your demo account, tap a tool, and borrow it in one click.
* **Google Sheets Backend:** Uses a connected Google Sheet as a live cloud database, allowing organizers to manage inventory or test live sync during demos effortlessly.
* **Instant Search & Filter:** Quickly filter through tools by name, category or neighbor.
* **Photo Listings:** Marketplace-style cards with a photo for every tool.
* **Reservations:** A two-week availability calendar; book dates ahead and borrow dates are capped so they never run into someone else's booking.
* **Waitlist:** Join the line for a checked-out tool; returns are held for whoever is next.
* **Friendly Nudges:** One-tap WhatsApp/SMS reminders for overdue tools and "it's back" hand-offs.
* **Condition Reports:** Returns log Good / Needs attention / Broken; broken tools go out for repair until the owner marks them fixed.
* **Trust Badges:** Every neighbor gets a track record (Trusted, Reliable, Often late) from their on-time return rate.
* **Wish List:** Ask the street for a tool you need; anyone who has one can list it straight from the request.
* **Community Impact:** Money saved, CO₂ and landfill waste avoided, top sharers, and most-borrowed tools.
* **Guided Tour:** A "Take the tour" button walks through the whole pitch on a sandboxed copy of the data.
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
├── assets/
│   ├── favico.png      # Navbar logo
│   ├── image.png       # Browser-tab favicon
│   └── tools/          # Sample tool photos (Wikimedia Commons, see CREDITS.md)
├── specsheet.txt       # Hackathon spec
└── README.md           # Project documentation
```

---

## Quick Start

1. Open `index.html` in a browser. With no backend configured it runs on demo data.
2. To go live: create a Google Sheet with an `Inventory` tab (headers: ID, Tool Name, Category, Status, Borrower, Due Date, Photo, Owner, Waitlist, Reservations, Condition, Condition Note). `Loans` and `Requests` tabs are created automatically.
3. Extensions → Apps Script → paste `apps-script/Code.gs` → Deploy as a Web app (Execute as: Me, Access: Anyone).
4. Paste the `/exec` URL into `SHEET_ENDPOINT` at the top of the script in `index.html`.
