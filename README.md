# Azure DevOps Hub

A fully functional **Azure DevOps Centralized Dashboard** built with plain HTML, CSS, and vanilla JavaScript (no framework, no build tools). Connect to multiple Azure DevOps organizations using a URL and Personal Access Token (PAT), and view all your projects, work items, pipelines, and dashboards in one place.

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🔗 **Connection Manager** | Add, edit, delete and toggle Azure DevOps connections. PAT is masked in the UI. Connections are validated before saving. |
| 📋 **Projects Listing** | Card-based grid showing all projects across all connected organizations, with search and org filter. |
| 🔧 **Work Items** | Table view of work items with filters (type, state, priority, search), pagination (50/page), and a detail modal. |
| 🚀 **Pipelines / Builds** | List of build pipelines with color-coded status badges, branch, triggered-by, duration and last run time. |
| 📊 **Dashboard** | Aggregate stats (orgs, projects, builds), build health bar chart, organization health cards, and recent activity feed. |
| 🌗 **Dark / Light Theme** | Toggle between a light and dark theme. Preference is persisted in `localStorage`. |
| 📱 **Responsive** | Works on desktop and tablet screens. |

---

## 📁 File Structure

```
index.html          ← SPA shell (sidebar + header + main area)
css/
  styles.css        ← All styling (theme tokens, layout, components)
js/
  api.js            ← Azure DevOps REST API wrapper (fetch + auth + errors)
  connections.js    ← Connection manager (localStorage CRUD + UI)
  projects.js       ← Projects listing page
  workitems.js      ← Work items page
  pipelines.js      ← Pipelines page
  dashboard.js      ← Dashboard page
  app.js            ← SPA router, global events, toasts, modal
README.md
```

---

## 🚀 Getting Started

No build tools or server required.

1. **Clone or download** this repository.
2. Open `index.html` directly in your browser  
   *(e.g., double-click the file, or use a simple static server like `npx serve .`)*
3. Click **Connections** in the sidebar and add your first Azure DevOps organization.
4. Explore your projects, work items, and pipelines from the sidebar.

---

## 🔑 How to Generate an Azure DevOps PAT

1. Sign in to [dev.azure.com](https://dev.azure.com).
2. Click your profile icon (top right) → **Personal access tokens**.
3. Click **+ New Token**.
4. Give the token a name and select an expiry.
5. Under **Scopes**, choose:
   - **Work Items** → Read
   - **Build** → Read
   - **Release** → Read
   - **Project and Team** → Read
6. Click **Create** and **copy the token immediately** (it's only shown once).
7. Paste the token into the PAT field when adding a connection in Azure DevOps Hub.

> ⚠️ **Security note:** The PAT is stored in your browser's `localStorage`. Do not use this app on shared or public computers.

---

## ⚠️ CORS Limitations

Azure DevOps REST APIs do **not** set CORS headers that allow arbitrary browser origins. This means that when you open `index.html` directly from the filesystem (`file://`) or from a different domain, API calls may be blocked by the browser.

**Workarounds:**

| Option | Description |
|--------|-------------|
| **Local dev server** | `npx serve .` or `python -m http.server` — still blocked by Azure DevOps CORS policy |
| **Browser extension** | Use a CORS-unblocking extension (development only) |
| **Proxy** | Set up a simple reverse proxy (e.g., nginx, Azure API Management) that adds CORS headers |
| **Electron / desktop wrapper** | Wrap the app in Electron to bypass browser CORS restrictions |

The app will display a helpful message when a CORS error is detected.

---

## 🖥️ Screenshots

> *Add screenshots here once the app is running.*

| Dashboard | Projects | Work Items | Pipelines | Connections |
|-----------|----------|------------|-----------|-------------|
| _(screenshot)_ | _(screenshot)_ | _(screenshot)_ | _(screenshot)_ | _(screenshot)_ |

---

## 🤝 Contributing

1. Fork the repository.
2. Create a feature branch: `git checkout -b feature/my-feature`.
3. Make your changes (keep it plain HTML/CSS/JS — no build tools).
4. Test by opening `index.html` in a browser.
5. Open a Pull Request with a description of your changes.

---

## 📄 License

MIT — see [LICENSE](LICENSE) for details.