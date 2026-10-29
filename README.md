# 🎓 Student Record Management System (College CA2 Project)

A complete, clean, and beginner-friendly web application for managing student academic records, built using **Python**, **Flask**, **MongoDB**, **PyMongo**, **HTML5**, and **CSS3**.

---

## 📌 Project Overview

This project was developed for the **College Continuous Assessment (CA2)**. It provides a simple, modern, and intuitive interface for college administrators and teachers to perform full **CRUD (Create, Read, Update, Delete)** operations on student records stored in a NoSQL MongoDB database.
🚀 **Live Demo:** [Open Project](https://student-record-management-1-nsq4.onrender.com)

### ✨ Key Features

1. **Dashboard & Summary Analytics**:
   - Total Students enrolled.
   - Average CGPA across all students.
   - Highest CGPA achieved.
   - Active courses counter.
   - Live MongoDB server connectivity indicator (`Connected` / `Offline`).

2. **Add Student (Create)**:
   - Enter student details: **Full Name**, **Registration Number**, **Course**, **Semester**, **College Email**, and **CGPA**.
   - Input validation: Email format, semester bounds (1–8), and CGPA range (0.00 to 10.00).
   - **Duplicate Prevention**: Rejects duplicate registration numbers at both the application level and via a MongoDB unique index.
   - Preserves user input if validation or duplicate checks fail.

3. **View Students (Read)**:
   - Displays all records in a responsive, clean table.
   - Formatted badges for Registration Number, Course, and Semester.
   - Color-coded CGPA pills (Outstanding >= 8.5, Good >= 7.0, Average < 7.0).

4. **Search & Filter**:
   - Real-time search by **Registration Number** or **Student Name**.
   - Filter dropdown by **Course / Program**.
   - One-click **Clear Filters** reset button.

5. **Update Student (Update)**:
   - Edit any existing student details via a pre-filled edit form.
   - Prevents registration number collision with other students.

6. **Delete Student (Delete)**:
   - One-click deletion with a safety confirmation prompt (`JavaScript confirm()`) to prevent accidental data loss.
   - Displays a success confirmation message showing the deleted student's name and registration number.

7. **Demo Data Generator**:
   - Built-in one-click **"Load Demo Data"** button in the web navigation bar.
   - Standalone CLI script (`seed_data.py`) to insert sample students instantly.

---

## 🛠️ Technology Stack

| Component | Technology | Description |
| :--- | :--- | :--- |
| **Backend Framework** | Python 3 + Flask | Lightweight WSGI web application framework |
| **Database** | MongoDB (Community Server / Atlas) | Document-oriented NoSQL database |
| **Database Driver** | PyMongo | Official Python driver for MongoDB |
| **Environment Config** | Python-Dotenv | Secure environment variable management (`.env`) |
| **Frontend** | HTML5 & Vanilla CSS3 | Responsive UI with Google Font (*Plus Jakarta Sans*) |

---

## 📁 Project Directory Structure

```text
student-record-management/
├── app.py                  # Main Flask application & CRUD route handlers
├── requirements.txt        # Python package dependencies
├── .env                    # Local environment variables (active config)
├── .env.example            # Template for environment variables
├── .gitignore              # Files to ignore in Git version control
├── seed_data.py            # CLI script to insert sample student records
├── test_app.py             # Automated pytest suite (7 CRUD & validation tests)
├── README.md               # Project documentation & presentation guide
├── templates/
│   ├── index.html          # Dashboard, stats, search, & student table
│   ├── add_student.html    # Form to add a new student record
│   └── edit_student.html   # Form to edit an existing student record
└── static/
    └── style.css           # Custom responsive stylesheet (design system)
```

---

## ⚙️ Step-by-Step Setup & Installation Guide

### Prerequisites
- **Python 3.8+** installed ([Download Python](https://www.python.org/downloads/))
- **MongoDB Community Server** installed and running locally, OR a free **MongoDB Atlas** cloud connection string.

---

### Step 1: Clone or Navigate to Project Folder
Open your terminal (PowerShell, Command Prompt, or Bash) in the project directory:
```bash
cd "c:\Users\Nirmal Mehra\Desktop\mondoDB"
```

---

### Step 2: (Optional) Create & Activate a Virtual Environment
It is recommended to use a virtual environment:

**On Windows (PowerShell):**
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

**On macOS / Linux:**
```bash
python3 -m venv venv
source venv/bin/activate
```

---

### Step 3: Install Required Dependencies
Run:
```bash
pip install -r requirements.txt
```
This installs:
- `Flask`
- `pymongo`
- `python-dotenv`

---

### Step 4: Configure MongoDB Connection
Check the `.env` file in the root folder. If you are using local MongoDB:
```env
SECRET_KEY=college_ca2_student_record_secret_key_2026
FLASK_DEBUG=True
PORT=5000

# Local MongoDB default connection
MONGO_URI=mongodb://localhost:27017/
DB_NAME=student_records_db
COLLECTION_NAME=students
```

> **Using MongoDB Atlas (Cloud)?**
> Replace `MONGO_URI` with your connection string:
> `MONGO_URI=mongodb+srv://<username>:<password>@cluster0.mongodb.net/?retryWrites=true&w=majority`

---

### Step 5: (Optional) Insert Demo Student Records
To populate the database with realistic sample students, run:
```bash
python seed_data.py
```
*(You can also simply click the **"⚡ Load Demo Data"** button in the web navigation bar after starting the app!)*

---

### Step 6: Start the Flask Application
Run:
```bash
python app.py
```

You should see output similar to:
```text
============================================================
>> Student Record Management System - Flask & MongoDB
============================================================
Server running at: http://127.0.0.1:5000
MongoDB URI:       mongodb://localhost:27017/
Database:          student_records_db
Collection:        students
============================================================
 * Running on http://127.0.0.1:5000
```

Open your browser and navigate to:
👉 **[http://127.0.0.1:5000](http://127.0.0.1:5000)**

---

## 🧪 Running Automated Tests

To run the built-in integration and CRUD test suite:
```bash
pytest test_app.py -v
```

This tests:
- Home page loading and dashboard metrics.
- Searching by registration number.
- Searching by student name.
- Adding a new student.
- Duplicate registration number prevention.
- Invalid CGPA bounds rejection.
- Full Create-Read-Update-Delete lifecycle.

---

## 🗄️ Database Schema & PyMongo Operations

Each student record is stored as a document in MongoDB:

```json
{
  "_id": ObjectId("670c79..."),
  "name": "Aarav Sharma",
  "reg_no": "2024BCSE001",
  "course": "B.Tech CSE",
  "semester": 6,
  "email": "aarav.sharma@college.edu",
  "cgpa": 9.24,
  "created_at": "2026-10-14T14:10:20.000Z",
  "updated_at": "2026-10-14T14:15:30.000Z"
}
```

### PyMongo Operations Implemented in `app.py`:

| Operation | PyMongo Method | Purpose in Project |
| :--- | :--- | :--- |
| **Unique Index** | `collection.create_index([("reg_no", 1)], unique=True)` | Ensures `reg_no` uniqueness at the database level |
| **Insert (Create)** | `collection.insert_one(student_document)` | Adds a new student record |
| **Find All (Read)** | `collection.find(query_filter).sort(...)` | Fetches student records for table and dashboard |
| **Find One (Search)**| `collection.find_one({"reg_no": reg_no})` | Searches for existing student or duplicate checks |
| **Find by ID** | `collection.find_one({"_id": ObjectId(id)})` | Retrieves a specific student to populate the edit form |
| **Update** | `collection.update_one({"_id": ...}, {"$set": ...})` | Updates modified student fields |
| **Delete** | `collection.delete_one({"_id": ObjectId(id)})` | Permanently deletes a student record |

---

## 🎓 College Viva / Presentation Q&A Guide

Prepare for your CA2 presentation and viva with these common questions:

### Q1: Why did you choose MongoDB instead of a traditional relational database (MySQL/PostgreSQL)?
> **Answer**: MongoDB is a document-oriented NoSQL database that stores data in flexible, JSON-like BSON documents. It does not require complex table joins and allows rapid schema evolution. For student management, where each student profile contains distinct academic and personal attributes, documents represent the real-world entity naturally.

### Q2: What is PyMongo and how does it connect Flask to MongoDB?
> **Answer**: PyMongo is the official Python distribution for working with MongoDB. In our Flask application, `pymongo.MongoClient()` initializes a connection pool to the MongoDB server using the `MONGO_URI`. Using `client[DB_NAME][COLLECTION_NAME]`, Flask routes execute queries like `insert_one()`, `find()`, `update_one()`, and `delete_one()`.

### Q3: How do you prevent duplicate registration numbers?
> **Answer**: Duplicate prevention is implemented at two layers:
> 1. **Application Layer**: Before inserting or updating, `collection.find_one({"reg_no": reg_no})` checks if the number is already taken.
> 2. **Database Layer**: A unique index `collection.create_index([("reg_no", 1)], unique=True)` is created on startup, which raises a `DuplicateKeyError` if any duplicate attempts bypass the application layer.

### Q4: What is `ObjectId` in MongoDB?
> **Answer**: Every document in MongoDB automatically receives a 12-byte hexadecimal unique identifier stored in the `_id` field. In Flask, we import `from bson.objectid import ObjectId` to convert the string ID from the URL (`/edit/<id>`) into an `ObjectId` object so MongoDB can locate the exact document.

### Q5: How are success and error notifications handled?
> **Answer**: We use Flask's built-in `flash()` messaging system with categories (`success`, `danger`, `warning`, `info`). In the HTML templates, Jinja2's `{% with messages = get_flashed_messages(with_categories=true) %}` dynamically renders styled alert boxes that users can dismiss.

---

## 📜 License
This project is open-source and intended for academic, college evaluation, and educational demonstration purposes.
