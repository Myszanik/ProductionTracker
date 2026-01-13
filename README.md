# ProductionTracker (prototype)

A Flask + SQLite production tracking prototype designed to replace paper tracking with a live overview of where orders are in production.

Built as a prototype for a friend, it was never deployed in a real business environment.

## What it does
- Separate logins for each station screen (to simulate different production areas)
- Stations can start and finish jobs, orders move through the workflow
- Capacity rules enforce realistic constraints (queue limits, wrapping slots, lorry capacity)
- Manager overview shows progression across stations
- Manager search shows the current location of an order quickly
- A history log tracks status changes over time

## Workflow
Preparing → CNC → Tramming 1 → Edge → Tramming 2 → Wrapping → Loading → Completed

CNC and Edge have multiple machines, the manager view groups them into single columns while still showing which machine was used.

## Tech stack
- Python
- Flask
- SQLite
- HTML (Jinja2 templates) + CSS

## Run locally
```bash
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
# .venv\Scripts\activate   # Windows PowerShell

pip install -r requirements.txt
python app.py

Open:
```bash
http://localhost:5050

## Demo logins (testing only)
This project was never deployed. The login accounts exist only to simulate separate station screens during local testing.
Manager:
- manager / manager123
Stations:
- preparing / prep123
- cnc1, cnc2, cnc3 / cnc123
- tramming1 / tram123
- edge1, edge2, edge3, edge4 / edge123
- tramming2 / tram123
- wrapping / wrap123
- loading / load123

## Configuration (optional)
This app uses a Flask secret key for sessions (logins). For a local prototype you can run without setting anything.
If you want to set it explicitly:
```bash
export FLASK_SECRET_KEY="change_me"
`
```bash
python app.py
See `.env.example` for the available variables.

## Work in progress
Some additional pages were started for future expansion and KPI reporting:
- admin
- analytics
- downtime
- training matrix
- weekly staffing
## AI assistance
See `AI_USAGE.md`.

## License
All rights reserved. Shared for viewing and evaluation purposes only.