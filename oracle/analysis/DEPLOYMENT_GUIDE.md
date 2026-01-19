# OCI Free Tier Deployment Guide (Podman)

This guide explains how to deploy your Research Hub project to **Oracle Cloud Infrastructure (OCI)** using Podman.

## Prerequisites
1. **OCI Compute Instance**: Use an **ARM Ampere A1** shape (4 OCPUs, 24GB RAM) for the best free performance.
2. **OS**: Oracle Linux 8 or 9 (standard on OCI).
3. **Open Port 80**: 
   - In OCI Console: Go to Virtual Cloud Network -> Security Lists -> Add Ingress Rule for Port 80.
   - On the VM: `sudo firewall-cmd --permanent --add-service=http && sudo firewall-cmd --reload`

## 1. Install Podman
Oracle Linux comes with Podman pre-installed, but ensure it's up to date:
```bash
sudo dnf install -y podman
```

## 2. Prepare the Project
Clone your repository to the server:
```bash
git clone <your-repo-url>
cd project-rh/oracle/analysis
```

## 3. Run the Deployment Script
We have created a script to automate the setup:
```bash
chmod +x deployment-setup.sh
./deployment-setup.sh
```

## 4. Scalability Features Explained
*   **Gunicorn + Uvicorn**: Your FastAPI app now runs with 4 worker processes. This allows it to handle multiple requests simultaneously without blocking.
*   **Nginx Buffering**: Nginx sits in front of the app to handle slow clients, compress static assets (Gzip), and manage the load.
*   **Systemd**: If the instance reboots or a container crashes, Podman will use `systemd` to restart the Pod automatically.

## 5. Maintenance Commands
*   **View Logs**: `podman logs research-hub-app-run`
*   **Check Pod Status**: `podman pod ps`
*   **Restart Services**: `systemctl --user restart pod-research-hub-pod.service`
*   **Update App**:
    1. Update code: `git pull`
    2. Rebuild: `podman build -t research-hub-app .`
    3. Restart: `podman stop research-hub-app-run && podman start research-hub-app-run`

## 6. Security Note (SSL/HTTPS)
For production (port 443), you should use **Certbot** (Let's Encrypt). Once you have a domain, you can mount the certificates into the Nginx container and update `nginx.conf` to use port 443.
