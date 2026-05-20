# DCC Survey Appointment Map (Google Maps POC)

## Development plan (brief)
1. Build a static file structure (`index.html`, `style.css`, `app.js`, `data.json`) with no build tooling.
2. Implement map initialization and JSON data loading.
3. Add markers, clustering, info window details, and status-based marker visuals with surveyor initials.
4. Implement search, filters, clear/reset, refresh, summary counts, proximity list, and simple route demo.
5. Validate the app is static-host friendly and document usage and SPFx migration path.

## How to run locally
This is a static website with no build step.

### Option A: VS Code Live Server
1. Open this folder in VS Code.
2. Right-click `index.html` and choose **Open with Live Server**.

### Option B: Python local server
```bash
python3 -m http.server 8080
```
Then browse to: `http://localhost:8080`

> Note: `data.json` is loaded via `fetch`, so open through a local web server (not `file://`).

## Google Maps API key setup
1. Open `index.html`.
2. Find the Google script tag near the bottom.
3. Replace `YOUR_GOOGLE_MAPS_API_KEY` with your key.

```html
<script async defer src="https://maps.googleapis.com/maps/api/js?key=YOUR_GOOGLE_MAPS_API_KEY&callback=initMap"></script>
```

Keep real keys out of source control.

## Google APIs to enable
- **Maps JavaScript API** (required)
- **Directions API** (required for route demo button)

## What this POC demonstrates
- Interactive Google Map centered on Dublin.
- Local JSON sample data representing SharePoint-style survey appointments.
- Marker rendering with status colors and surveyor initials.
- Pending Allocation markers using `?` visual indicator.
- Marker click info window with appointment/property details.
- Search + status/surveyor filtering without page reload.
- Marker clustering for dense map areas.
- Summary counters (total, visible, status counts).
- Data refresh button and optional 60-second auto-refresh toggle in code.
- Nearby appointment proximity list (straight-line distance).
- Simple two-point route rendering via Directions Service.

## What is excluded from this POC
- SharePoint list integration (REST/Graph/SPFx data layer)
- SMS notifications
- Microsoft Bookings integration
- Power Automate flows
- Authentication/authorization and production hardening

## SPFx conversion approach (later)
1. Move map UI into an SPFx Web Part component.
2. Replace `fetch('./data.json')` with SharePoint list retrieval (REST or Graph).
3. Store API key securely (tenant config / Azure Key Vault pattern), not in page source.
4. Keep current filtering/rendering modules as plain TypeScript utility layers.
5. Add tenant-specific permissions, logging, and deployment pipeline controls.
