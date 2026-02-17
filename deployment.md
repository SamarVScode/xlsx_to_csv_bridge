# Deployment Specifications for Render.com

## Service Type
- **Web Service**

## Environment
- **Runtime**: Python 3
- **Branch**: main (or your repository branch)
- **Root Directory**: `.` (current directory)

## Build & Start Commands
| Configuration | Value |
| :--- | :--- |
| **Build Command** | `pip install -r requirements.txt` |
| **Start Command** | `uvicorn main:app --host 0.0.0.0 --port 10000` |

> [!NOTE]
> Render automatically sets the `PORT` environment variable to `10000` (or similar). `uvicorn` must listen on `0.0.0.0` to be accessible.

## Environment Variables
You MUST set the following environment variable in the Render Dashboard (Environment > Environment Variables):

| Key | Value | Description |
| :--- | :--- | :--- |
| `MY_SECRET_API_KEY` | `(your_secret_key_here)` | Security key matching the one in `Code.gs`. |
| `PYTHON_VERSION` | `3.9.0` | (Optional) Specify Python version if needed. |

## Plan
- **Instance Type**: Free
- **RAM**: 512 MB (This app is optimized for this limit via stream processing)
