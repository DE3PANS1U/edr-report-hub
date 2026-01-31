# Deployment Guide: EDR Report Hub

This guide will help you host your EDR Report Generator live on the web so anyone can access it. We recommend using **Render.com** as it offers a free tier and works seamlessly with the files we've prepared.

---

## 🚀 Quick Deployment to Render.com

### Prerequisites
1.  **GitHub Account**: You need a GitHub account to host your code.
2.  **Render Account**: Sign up at [dashboard.render.com](https://dashboard.render.com).

### Step 1: Push Code to GitHub
1.  Create a new repository on GitHub (e.g., `edr-report-hub`).
2.  Upload all files from your project folder to this repository.
    *   Ensure `requirements.txt`, `Procfile`, `app.py`, `edr_report_generator_custom.py`, `templates/`, and `static/` are included.

### Step 2: Create Web Service on Render
1.  Log in to your Render dashboard.
2.  Click **"New +"** and select **"Web Service"**.
3.  Connect your GitHub account and select the `edr-report-hub` repository you just created.
4.  Configure the service:
    *   **Name**: `edr-report-hub` (or your preferred name)
    *   **Region**: Closest to you (e.g., Singapore, Frankfurt)
    *   **Branch**: `main` (or `master`)
    *   **Runtime**: `Python 3`
    *   **Build Command**: `pip install -r requirements.txt`
    *   **Start Command**: `gunicorn app:app` (This should auto-fill thanks to the `Procfile`)
    *   **Instance Type**: `Free`

### Step 3: Go Live!
1.  Click **"Create Web Service"**.
2.  Render will start building your application. This may take a few minutes.
3.  Once the status says **"Live"**, you will see a URL (e.g., `https://edr-report-hub.onrender.com`).
4.  Share this URL with anyone who needs to access the tool!

---

## 📂 Deployment Files Explained

*   **`Procfile`**: Tells the cloud server how to start your app (`gunicorn app:app`).
*   **`requirements.txt`**: Lists all Python libraries needed (Flask, pandas, etc.).
*   **`app.py`**: The main application server.

---

## ⚠️ Important Notes
*   **Free Tier Limitations**: On the free tier, the "spin down" might occur after inactivity. The first request after a while might take 30-60 seconds to load. This is normal.
*   **File Storage**: Uploaded files and generated reports are stored temporarily. Render's free tier filesystem is ephemeral, meaning files are deleted when the app restarts (which is perfect for security).

---

## 🛡️ Support
Created by **Deepanshu Patel**
For issues, check the "Events" log in your Render dashboard.
