# Deployment Notes

This repository is the single source of truth for `/opt/reportgen-web` on both
the JD Cloud server and the VPN server.

Do not commit runtime data or secrets:

- `backend/.env`
- `backend/.admin_credentials`
- `config/patient_info.yaml`
- `storage/`
- generated reports, uploads, caches, and virtual environments

Bootstrap or refresh a server:

```bash
cd /opt/reportgen-web
git pull origin main
bash deploy.sh
```

For a first-time install, run:

```bash
REPO_URL=https://github.com/youngfly93/reportgen-web.git bash deploy.sh
```

The deploy script preserves existing `backend/.env` secrets when present, builds
the frontend, copies it to `backend/static`, installs Python dependencies, and
restarts `reportgen-web`.
