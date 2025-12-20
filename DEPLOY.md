# Deployment Instructions

## Environments

| Environment | Manifest | Hosted At |
|-------------|----------|-----------|
| **DEV** | `manifest.xml` (this folder) | `localhost:3000` |
| **PROD** | `../Task Manager Excel/prod/manifest.xml` | `https://elvince01.github.io/lcmc-task-manager-addin/` |

## Local Development

```bash
npm run dev-server   # Starts https://localhost:3000
```

Then sideload `manifest.xml` in Excel.

## Build & Deploy to Production

### Step 1: Build
```bash
npm run build
```

### Step 2: Copy to prod folder
```bash
cp dist/taskpane.html dist/taskpane.js dist/*.js "../../Task Manager Excel/prod/"
cp -r dist/assets "../../Task Manager Excel/prod/"
```

### Step 3: Push to GitHub Pages
```bash
cd "../../Task Manager Excel/prod"
# Copy contents to your GitHub Pages repo and push
# Repo: https://github.com/elvince01/lcmc-task-manager-addin
```

## Quick Deploy Script

```bash
npm run build && cp dist/taskpane.html dist/taskpane.js dist/*.js "../../Task Manager Excel/prod/"
```
