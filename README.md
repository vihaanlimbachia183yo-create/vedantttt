# Mumbai MHT CET CAP Option Form Planner

A production-oriented React/Vite web application for preparing an engineering CAP preference list for Mumbai City, Mumbai Suburban, Navi Mumbai, and optionally Thane.

## Data policy

The application intentionally ships with **no fabricated cutoff, institute, branch, fee, placement, or seat data**. Import official Maharashtra State CET Cell / official college records before use. If a field is not present in an official source, keep it blank and the UI displays `Official data unavailable`.

Supported import formats: CSV, Excel (`.xlsx`, `.xls`), and JSON. Required planning identifiers are college, branch, CAP code, and district. Recommended official audit columns:

- admissionYear
- capRound
- college
- branch
- capCode
- category
- seatType
- gender
- homeUniversity
- tfws
- minority
- closingPercentile
- sourceDocument
- sourcePage
- verificationStatus
- district
- instituteType
- autonomous
- fees
- intake
- placements
- website
- notes

District values are restricted to `Mumbai City`, `Mumbai Suburban`, `Navi Mumbai`, and `Thane`.

## Features

- Local student profile with CET percentile, category, gender, home university, seat type, minority, and TFWS eligibility.
- Import pipeline for official CSV/Excel/JSON datasets with validation and import history.
- Search by college, branch, and CAP code.
- Simultaneous filters for district, branch, category, and recommendation category.
- Recommendation engine: Dream, Reach, Competitive, Safe, Very Safe, including cutoff difference and probability band.
- CAP preference form with duplicate prevention, add/remove, move up/down, auto numbering, notes, autosave, and restore from local storage.
- Dashboard statistics for total preferences, category counts, highest/lowest/average cutoff.
- College details modal for official metadata.
- Export preferences as JSON, Excel, and printable PDF/browser print sheet.
- Dark mode and responsive layout.

## Installation

```bash
npm install
```

## Development

```bash
npm run dev
```

Open the Vite URL printed in the terminal.

## Production build

```bash
npm run build
npm run preview
```

## Testing

```bash
npm test
```

## Importing official CET files

1. Download the official participating institute list, seat matrix, CAP round cutoffs, and information brochure from Maharashtra State CET Cell or official college websites.
2. Convert PDFs to CSV/Excel using your audited workflow.
3. Preserve source document name and source page for every imported cutoff row.
4. Import the resulting CSV/XLSX/JSON file from the **Import** button.
5. Review rejected malformed records and fix the source spreadsheet rather than guessing missing values.

## Architecture

This compact Vite implementation keeps all runtime code client-side for deployability on static hosts while using normalized record shapes and explicit source fields so it can be backed by an Express/Prisma API later without changing the UI workflow.

- `App.jsx` — application state, validation, recommendation engine, imports, exports, preference workflow, details modal.
- `index.css` — responsive and dark-mode styling.
- `tests/validation.test.mjs` — validation-oriented smoke tests.

## Deployment

Build with `npm run build` and deploy the `dist/` directory to Netlify, Vercel, Cloudflare Pages, or any static host.

## Troubleshooting

- **No colleges visible:** import an official Mumbai/Navi Mumbai/Thane dataset first.
- **Rows rejected:** check district spelling and required identifiers.
- **Blank cutoffs:** official source did not provide a verified percentile for that row; do not estimate.
- **Exports missing data:** the preference form exports only rows you added to preferences.

## Screenshots

Run the app locally and capture screenshots after importing your official dataset. No screenshots are bundled because bundled data would risk presenting unaudited official values.
