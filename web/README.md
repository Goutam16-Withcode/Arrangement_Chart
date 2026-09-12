# DeskMatrix™ Web Application (Next.js 16 + React 19)

Frontend interface for **DeskMatrix™ Enterprise Exam Seating & Attendance System**.

---

## ⚡ Quick Start

```bash
# Install packages
npm install

# Start Next.js local server
npm run dev
```

The web studio will start on **[http://localhost:3000](http://localhost:3000)**.

---

## 🔌 Connecting to Python Backend

1. Set the backend URL in `.env.local`:
   ```env
   NEXT_PUBLIC_PYTHON_BACKEND_URL=http://127.0.0.1:8000
   ```
2. Start the Python FastAPI backend from the root directory:
   ```bash
   cd ..
   python api_server.py
   ```

For full documentation and screenshots, refer to the root [README.md](../README.md).
