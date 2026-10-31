# SeaZen Power Apps Simulator (`Pa_Simulator`)

Local Power Apps `.pa.yaml` Studio, Live Runtime Player, and SharePoint Online OAuth2 CRUD Engine (`640 x 1136` canvas).

## Core Capabilities

1. **Open Local Files & Packages (`File > Open`)**
   - Load any `.pa.yaml`, `.yaml`, or `.msapp` package directly from the local filesystem via native Windows File/Folder dialogs (`POST /api/dialog/open-file`, `POST /api/dialog/open-folder`).
   - Automatically extracts `.msapp` packages and parses the `Src/*.pa.yaml` control hierarchy.

2. **OAuth2 Bearer Token Authentication (No-Cookie Standard)**
   - Authenticates against Microsoft Graph API (`https://graph.microsoft.com/v1.0`) using SeaZen Azure App Registration (`SEAZENHR-PORTAL`: `0390280b-12b8-49c4-adbf-a9c2b8bccfb5` in Tenant `29101901-3778-4339-827c-7d2b50914db0`).
   - Supports silent token renewal via `token_cache.bin` and interactive OAuth2 login (`POST /api/auth/login`).

3. **Live SharePoint CRUD & Power Fx Execution**
   - Renders Power Apps controls (`GroupContainer`, `Gallery`, `FluidGrid`, `DataCard`, `HtmlViewer`, `Classic/Button`, `Label`, `Classic/TextInput`, `Classic/Dropdown`, `DatePicker`, `Classic/Icon`, `Rectangle`).
   - Executes Power Fx formulas (`Patch`, `Remove`, `Navigate`, `Back`, `Set`, `UpdateContext`, `Notify`) directly against SharePoint Online lists (`POST /api/sp/crud`).
   - Built-in **SharePoint Live CRUD Console** to inspect, create (`POST`), update (`PATCH`), and delete (`DELETE`) items across SharePoint Online lists.
   - Static linter for SharePoint 5,000-row delegation limits and missing `ThisItem` properties.

4. **Standalone Published Player Mode**
   - Run in Studio mode (`http://localhost:8095`) or Published Player mode (`http://localhost:8095/player?mode=published`).

## Quick Start

```bash
pip install -r requirements.txt
python sandbox_server.py
```

Or double-click:
- `Start_Studio.bat` — Launches the 3-pane YAML Studio & Live Preview (`http://localhost:8095`).
- `Start_Published_Player.bat` — Launches the standalone Published App Player (`http://localhost:8095/player?mode=published`).
