# EasySource HRMS Enterprise 🏢

Full-featured HR Management System with WhatsApp Messaging

## 🚀 Quick Start

### Windows (Double-click)
```
start_server.bat
```
Then open: **http://localhost:5000**

### Linux / Mac / VPS
```bash
pip install flask
python app.py
```

## 🔑 Default Login

| Role | Username | Password |
|------|----------|----------|
| Admin | admin | admin123 |
| User | user1 | user123 |

---

## 📦 Features

### 👨‍💼 Employees
- Add / Edit / Delete employees
- Card view + List view
- Search & filter
- **Bulk upload via Excel** (.xlsx, .csv)
- Departments & designations

### 📅 Attendance
- Mark Present / Absent / On Duty (OD)
- Add remarks per employee
- Load by date
- **Bulk mark all present** with one click
- Auto-saves when switching dates

### 📆 Weekly Roster
- Work / Weekly Off / Holiday / Leave
- Navigate weeks forward/backward
- Color-coded day status
- Save entire week's roster at once

### 📲 WhatsApp Messaging
- Select WhatsApp groups
- Send immediate messages
- **Schedule messages** (Admin/SubAdmin)
- **Auto-send daily** at set time (Admin)
- Message variables: {date}, {time}, {company}
- Full message history

### 👥 User Roles

| Feature | Admin | SubAdmin | User | Viewer |
|---------|-------|----------|------|--------|
| Dashboard | ✅ | ✅ | ✅ | ✅ |
| Attendance | ✅ | ✅ | ✅ | ❌ |
| Add Employees | ✅ | ✅ | ❌ | ❌ |
| Delete Employees | ✅ | ❌ | ❌ | ❌ |
| Send WhatsApp | ✅ | ✅ | ✅ | ❌ |
| Schedule WA | ✅ | ✅ | ❌ | ❌ |
| Manage Roster | ✅ | ✅ | ❌ | ❌ |
| Reports + Export | ✅ | ✅ | ✅ | ✅ |
| User Management | ✅ | ❌ | ❌ | ❌ |

### 📊 Reports
- Filter by date range
- Filter by department
- Summary statistics
- **Export to Excel** (.xlsx)

---

## 📲 WhatsApp API Setup

### Option 1: UltraMsg (Recommended – Easy)
1. Go to https://ultramsg.com
2. Create account → Add instance → Scan QR
3. Copy Instance ID and Token
4. Paste in HRMS Settings → WhatsApp API Config

### Option 2: WA-API.com
1. Go to https://wa-api.com
2. Same process as UltraMsg

### Option 3: Meta WhatsApp Cloud API (Official)
1. Go to https://developers.facebook.com
2. Create app → WhatsApp → Get token and Phone Number ID

**In HRMS:** Settings → WhatsApp API Config → Enter credentials → Save

---

## 🌐 VPS / Server Deployment

### On VPS (Ubuntu)
```bash
# Install
sudo apt update && sudo apt install python3 python3-pip -y
pip3 install flask gunicorn

# Run with Gunicorn (production)
gunicorn -w 4 -b 0.0.0.0:5000 app:app

# Or with Nginx reverse proxy for domain
# Configure Nginx to proxy to localhost:5000
```

### With Domain
Point your domain → VPS IP → Set up Nginx → SSL with Let's Encrypt

---

## 📁 File Structure
```
HRMS/
├── app.py              ← Main application
├── requirements.txt    ← Dependencies
├── start_server.bat    ← Windows start script
├── database/
│   └── hrms.db        ← SQLite database (auto-created)
└── templates/
    ├── login.html
    ├── base.html
    ├── dashboard.html
    ├── employees.html
    ├── attendance.html
    ├── roster.html
    ├── whatsapp.html
    ├── reports.html
    └── settings.html
```

## 🔄 Excel Upload Format

| emp_id | name | phone | department | designation | email | join_date |
|--------|------|-------|------------|-------------|-------|-----------|
| E001 | Rahul Kumar | 9876543210 | Operations | Manager | r@co.com | 2024-01-15 |

---

**Made with ❤️ for EasySource | Enterprise HRMS**
