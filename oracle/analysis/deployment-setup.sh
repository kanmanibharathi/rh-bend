#!/bin/bash

# --- Oracle Cloud OCI Deployment Script (Podman) ---
# This script sets up the backend (FastAPI) and frontend (Nginx) in a single Pod.

set -e

# Configuration
POD_NAME="research-hub-pod"
APP_IMAGE="research-hub-app"
NGINX_CONF="$(pwd)/nginx.conf"
FRONTEND_DIR="$(pwd)/../../cloudflare"

echo "[1/5] Building Python Backend Image..."
podman build -t $APP_IMAGE -f Containerfile .

echo "[2/5] Creating Pod..."
# Map port 80 (Host) to port 80 (Pod)
if ! podman pod exists $POD_NAME; then
    podman pod create --name $POD_NAME -p 80:80
else
    echo "Pod $POD_NAME already exists."
fi

echo "[3/5] Starting Backend Container..."
# Run inside the pod. It will be accessible at localhost:8000 within the pod.
podman run -d \
    --pod $POD_NAME \
    --name research-hub-app-run \
    --restart always \
    $APP_IMAGE

echo "[4/5] Starting Nginx Reverse Proxy..."
# Run inside the pod. It maps port 80 of the pod to its own port 80.
# We mount the nginx.conf and the frontend static files.
podman run -d \
    --pod $POD_NAME \
    --name research-hub-nginx \
    --restart always \
    -v "$NGINX_CONF:/etc/nginx/nginx.conf:ro" \
    -v "$FRONTEND_DIR:/usr/share/nginx/html:ro" \
    nginx:alpine

echo "[5/5] Generating Systemd Units for Auto-start..."
# This ensures the pod and containers start automatically after a reboot.
mkdir -p ~/.config/systemd/user
cd ~/.config/systemd/user
podman generate systemd --new --name $POD_NAME --files

systemctl --user daemon-reload
systemctl --user enable pod-$POD_NAME.service

echo "------------------------------------------------"
echo "Deployment Complete!"
echo "Your app is now running on port 80."
echo "Auto-start is enabled via systemd."
echo "Check status with: podman pod ps"
echo "------------------------------------------------"
