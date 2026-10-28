"""
=============================================================================
Student Record Management System
College CA2 Project
Technologies: Python, Flask, MongoDB, PyMongo, HTML5, CSS3
=============================================================================
"""

import os
import re
from datetime import datetime
from bson.objectid import ObjectId
from dotenv import load_dotenv
from flask import Flask, render_template, request, redirect, url_for, flash
import pymongo
from pymongo.errors import ConnectionFailure, DuplicateKeyError, PyMongoError

# Load environment variables from .env file
load_dotenv()

# Initialize Flask Application
app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "default_secret_key_ca2_student_system")

# MongoDB Configuration
MONGO_URI = os.getenv("MONGO_URI", "mongodb://localhost:27017/")
DB_NAME = os.getenv("DB_NAME", "student_records_db")
COLLECTION_NAME = os.getenv("COLLECTION_NAME", "students")

# Global Database Variables
mongo_client = None
db = None
students_collection = None
is_mock_db = False


def get_db():
    """
    Establish and return the MongoDB collection for students.
    Includes connection retry and sets a unique index on 'reg_no'.
    If a local MongoDB instance is not active, automatically falls back to
    mongomock so the application and all CRUD operations work seamlessly.
    """
    global mongo_client, db, students_collection, is_mock_db
    if students_collection is not None:
        return students_collection, None

    try:
        # Try real MongoDB connection first with a 1.5s timeout
        client = pymongo.MongoClient(MONGO_URI, serverSelectionTimeoutMS=1500)
        client.admin.command('ping')
        mongo_client = client
        db = mongo_client[DB_NAME]
        students_collection = db[COLLECTION_NAME]
        students_collection.create_index([("reg_no", pymongo.ASCENDING)], unique=True)
        is_mock_db = False
        return students_collection, None
    except (ConnectionFailure, PyMongoError, Exception) as conn_err:
        # Fall back to mongomock so the application remains completely functional
        try:
            import mongomock
            print(f">> [Notice] Local MongoDB server not reachable at {MONGO_URI}. Using in-memory MongoDB (mongomock).")
            mongo_client = mongomock.MongoClient()
            db = mongo_client[DB_NAME]
            students_collection = db[COLLECTION_NAME]
            students_collection.create_index([("reg_no", pymongo.ASCENDING)], unique=True)
            is_mock_db = True

            # Pre-seed demo records so tests and the dashboard have immediate data
            if students_collection.count_documents({}) == 0:
                for demo in DEMO_STUDENTS:
                    doc = dict(demo)
                    doc["created_at"] = datetime.now()
                    students_collection.insert_one(doc)

            return students_collection, None
        except Exception as mock_err:
            mongo_client = None
            db = None
            students_collection = None
            return None, f"Could not connect to MongoDB server: {str(conn_err)}"


# Sample demo students data for quick setup/evaluation
DEMO_STUDENTS = [
    {
        "name": "Aarav Sharma",
        "reg_no": "2024BCSE001",
        "course": "B.Tech CSE",
        "semester": 6,
        "email": "aarav.sharma@college.edu",
        "cgpa": 9.24,
        "created_at": datetime.now()
    },
    {
        "name": "Priya Patel",
        "reg_no": "2024BCSE042",
        "course": "B.Tech IT",
        "semester": 6,
        "email": "priya.patel@college.edu",
        "cgpa": 8.85,
        "created_at": datetime.now()
    },
    {
        "name": "Rohan Verma",
        "reg_no": "2024BCA015",
        "course": "BCA",
        "semester": 4,
        "email": "rohan.verma@college.edu",
        "cgpa": 7.92,
        "created_at": datetime.now()
    },
    {
        "name": "Sneha Kulkarni",
        "reg_no": "2024MCA008",
        "course": "MCA",
        "semester": 2,
        "email": "sneha.kulkarni@college.edu",
        "cgpa": 9.60,
        "created_at": datetime.now()
    },
    {
        "name": "Devansh Singh",
        "reg_no": "2024BECE029",
        "course": "B.Tech ECE",
        "semester": 4,
        "email": "devansh.singh@college.edu",
        "cgpa": 8.15,
        "created_at": datetime.now()
    }
]


# =============================================================================
# Helper Functions
# =============================================================================

def validate_student_data(name, reg_no, course, semester, email, cgpa):
    """
    Validates all input fields for student records.
    Returns (is_valid, error_message, cleaned_data)
    """
    # 1. Check for empty fields
    if not name or not reg_no or not course or not semester or not email or cgpa == "":
        return False, "All fields are required. Please fill in all the details.", None

    name = name.strip()
    reg_no = reg_no.strip().upper()
    course = course.strip()
    email = email.strip().lower()

    # 2. Validate Registration Number (letters, numbers, hyphens allowed)
    if len(reg_no) < 3 or len(reg_no) > 20:
        return False, "Registration Number must be between 3 and 20 characters.", None

    # 3. Validate Email format
    email_pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    if not re.match(email_pattern, email):
        return False, "Please provide a valid email address.", None

    # 4. Validate Semester (integer between 1 and 8)
    try:
        semester = int(semester)
        if semester < 1 or semester > 10:
            return False, "Semester must be a number between 1 and 10.", None
    except ValueError:
        return False, "Semester must be a valid integer.", None

    # 5. Validate CGPA (float between 0.00 and 10.00)
    try:
        cgpa = float(cgpa)
        if cgpa < 0.0 or cgpa > 10.0:
            return False, "CGPA must be between 0.00 and 10.00.", None
        cgpa = round(cgpa, 2)
    except ValueError:
        return False, "CGPA must be a valid decimal number (e.g. 8.5).", None

    cleaned_data = {
        "name": name,
        "reg_no": reg_no,
        "course": course,
        "semester": semester,
        "email": email,
        "cgpa": cgpa
    }
    return True, None, cleaned_data


# =============================================================================
# Flask Routes
# =============================================================================

@app.route("/")
def index():
    """
    Home page route:
    - Displays all student records in a table
    - Supports search by registration number or student name
    - Displays summary dashboard statistics (Total, Avg CGPA, Highest CGPA)
    """
    collection, error = get_db()
    if error:
        flash(error, "danger")
        return render_template(
            "index.html",
            students=[],
            stats={"total": 0, "avg_cgpa": 0, "highest_cgpa": 0, "courses_count": 0},
            search_query="",
            selected_course="",
            courses=[],
            db_connected=False
        )

    # Search and Filter Parameters
    search_query = request.args.get("search", "").strip()
    selected_course = request.args.get("course", "").strip()

    # Build MongoDB query filter
    query_filter = {}
    if search_query:
        # Search by registration number OR student name (case-insensitive regex)
        query_filter["$or"] = [
            {"reg_no": {"$regex": re.escape(search_query), "$options": "i"}},
            {"name": {"$regex": re.escape(search_query), "$options": "i"}}
        ]

    if selected_course:
        query_filter["course"] = selected_course

    try:
        # Retrieve filtered students sorted by creation time / registration number
        students_cursor = collection.find(query_filter).sort("created_at", pymongo.DESCENDING)
        students_list = list(students_cursor)

        # Retrieve distinct courses for the filter dropdown
        all_courses = sorted(collection.distinct("course"))

        # Calculate dashboard statistics
        all_students = list(collection.find({}, {"cgpa": 1, "course": 1}))
        total_students = len(all_students)
        
        if total_students > 0:
            avg_cgpa = round(sum(s.get("cgpa", 0.0) for s in all_students) / total_students, 2)
            highest_cgpa = max(s.get("cgpa", 0.0) for s in all_students)
            courses_count = len(set(s.get("course", "") for s in all_students if s.get("course")))
        else:
            avg_cgpa = 0.0
            highest_cgpa = 0.0
            courses_count = 0

        stats = {
            "total": total_students,
            "avg_cgpa": avg_cgpa,
            "highest_cgpa": highest_cgpa,
            "courses_count": courses_count
        }

        return render_template(
            "index.html",
            students=students_list,
            stats=stats,
            search_query=search_query,
            selected_course=selected_course,
            courses=all_courses,
            db_connected=True,
            is_mock=is_mock_db
        )

    except PyMongoError as pe:
        flash(f"Database error while fetching students: {str(pe)}", "danger")
        return render_template(
            "index.html",
            students=[],
            stats={"total": 0, "avg_cgpa": 0, "highest_cgpa": 0, "courses_count": 0},
            search_query="",
            selected_course="",
            courses=[],
            db_connected=False
        )


@app.route("/search")
def search():
    """
    Dedicated search route that forwards query parameter to index route.
    Satisfies Requirement #4: 'Search Student by registration number'.
    """
    query = request.args.get("q", "").strip()
    return redirect(url_for("index", search=query))


@app.route("/add", methods=["GET", "POST"])
def add_student():
    """
    Add Student route:
    - GET: Displays the Add Student form
    - POST: Validates input, prevents duplicate reg_no, inserts record into MongoDB
    """
    collection, error = get_db()
    if error:
        flash(error, "danger")
        return redirect(url_for("index"))

    if request.method == "POST":
        name = request.form.get("name", "")
        reg_no = request.form.get("reg_no", "")
        course = request.form.get("course", "")
        semester = request.form.get("semester", "")
        email = request.form.get("email", "")
        cgpa = request.form.get("cgpa", "")

        # Form validation
        is_valid, validation_error, cleaned_data = validate_student_data(
            name, reg_no, course, semester, email, cgpa
        )

        if not is_valid:
            flash(validation_error, "danger")
            # Preserve user input on validation error
            return render_template(
                "add_student.html",
                form_data=request.form
            )

        try:
            # Check if student with this registration number already exists
            existing_student = collection.find_one({"reg_no": cleaned_data["reg_no"]})
            if existing_student:
                flash(
                    f"Duplicate Entry: A student with Registration Number '{cleaned_data['reg_no']}' already exists ({existing_student.get('name')}).",
                    "warning"
                )
                return render_template("add_student.html", form_data=request.form)

            # Insert new record into MongoDB
            student_document = {
                "name": cleaned_data["name"],
                "reg_no": cleaned_data["reg_no"],
                "course": cleaned_data["course"],
                "semester": cleaned_data["semester"],
                "email": cleaned_data["email"],
                "cgpa": cleaned_data["cgpa"],
                "created_at": datetime.now()
            }

            collection.insert_one(student_document)
            flash(f"Student record for '{cleaned_data['name']}' ({cleaned_data['reg_no']}) added successfully!", "success")
            return redirect(url_for("index"))

        except DuplicateKeyError:
            flash(f"Error: A student with Registration Number '{cleaned_data['reg_no']}' already exists.", "warning")
            return render_template("add_student.html", form_data=request.form)
        except PyMongoError as pe:
            flash(f"Database error while adding student: {str(pe)}", "danger")
            return render_template("add_student.html", form_data=request.form)

    # GET request: render blank form
    return render_template("add_student.html", form_data={})


@app.route("/edit/<id>", methods=["GET", "POST"])
def edit_student(id):
    """
    Update Student route:
    - GET: Retrieves student record by ObjectId and pre-populates edit form
    - POST: Validates input, prevents duplicate reg_no conflict, updates document in MongoDB
    """
    collection, error = get_db()
    if error:
        flash(error, "danger")
        return redirect(url_for("index"))

    # Validate MongoDB ObjectId format
    try:
        student_obj_id = ObjectId(id)
    except Exception:
        flash("Invalid student identifier format.", "danger")
        return redirect(url_for("index"))

    # Fetch existing student
    student = collection.find_one({"_id": student_obj_id})
    if not student:
        flash("Student record not found.", "warning")
        return redirect(url_for("index"))

    if request.method == "POST":
        name = request.form.get("name", "")
        reg_no = request.form.get("reg_no", "")
        course = request.form.get("course", "")
        semester = request.form.get("semester", "")
        email = request.form.get("email", "")
        cgpa = request.form.get("cgpa", "")

        # Form validation
        is_valid, validation_error, cleaned_data = validate_student_data(
            name, reg_no, course, semester, email, cgpa
        )

        if not is_valid:
            flash(validation_error, "danger")
            # Preserve form data
            student_data = dict(student)
            student_data.update(request.form)
            return render_template("edit_student.html", student=student_data)

        try:
            # Check if updated registration number is already taken by ANOTHER student
            duplicate_check = collection.find_one({
                "_id": {"$ne": student_obj_id},
                "reg_no": cleaned_data["reg_no"]
            })
            if duplicate_check:
                flash(
                    f"Cannot update: Registration Number '{cleaned_data['reg_no']}' is already assigned to {duplicate_check.get('name')}.",
                    "warning"
                )
                student_data = dict(student)
                student_data.update(request.form)
                return render_template("edit_student.html", student=student_data)

            # Update document in MongoDB
            update_fields = {
                "name": cleaned_data["name"],
                "reg_no": cleaned_data["reg_no"],
                "course": cleaned_data["course"],
                "semester": cleaned_data["semester"],
                "email": cleaned_data["email"],
                "cgpa": cleaned_data["cgpa"],
                "updated_at": datetime.now()
            }

            collection.update_one({"_id": student_obj_id}, {"$set": update_fields})
            flash(f"Student record for '{cleaned_data['name']}' updated successfully!", "success")
            return redirect(url_for("index"))

        except DuplicateKeyError:
            flash(f"Error: Registration Number '{cleaned_data['reg_no']}' is already in use.", "warning")
            student_data = dict(student)
            student_data.update(request.form)
            return render_template("edit_student.html", student=student_data)
        except PyMongoError as pe:
            flash(f"Database error while updating student: {str(pe)}", "danger")
            student_data = dict(student)
            student_data.update(request.form)
            return render_template("edit_student.html", student=student_data)

    # GET request: render edit form with existing details
    return render_template("edit_student.html", student=student)


@app.route("/delete/<id>", methods=["POST", "GET"])
def delete_student(id):
    """
    Delete Student route:
    - Removes student record from MongoDB
    - Includes friendly flash confirmation message
    """
    collection, error = get_db()
    if error:
        flash(error, "danger")
        return redirect(url_for("index"))

    try:
        student_obj_id = ObjectId(id)
    except Exception:
        flash("Invalid student identifier format.", "danger")
        return redirect(url_for("index"))

    try:
        # Find record first to get student name for confirmation
        student = collection.find_one({"_id": student_obj_id})
        if not student:
            flash("Student record not found or already deleted.", "warning")
            return redirect(url_for("index"))

        # Delete from MongoDB
        collection.delete_one({"_id": student_obj_id})
        flash(f"Student record for '{student.get('name')}' ({student.get('reg_no')}) has been permanently deleted.", "success")
        return redirect(url_for("index"))

    except PyMongoError as pe:
        flash(f"Database error while deleting student: {str(pe)}", "danger")
        return redirect(url_for("index"))


@app.route("/load-demo-data")
def load_demo_data():
    """
    Utility route to seed sample student records into MongoDB.
    Useful for college viva, demonstrations, and initial testing.
    """
    collection, error = get_db()
    if error:
        flash(error, "danger")
        return redirect(url_for("index"))

    try:
        inserted_count = 0
        skipped_count = 0

        for demo in DEMO_STUDENTS:
            # Check if student with this reg_no already exists
            if not collection.find_one({"reg_no": demo["reg_no"]}):
                # Copy demo record to avoid mutating global template
                student_copy = dict(demo)
                student_copy["created_at"] = datetime.now()
                collection.insert_one(student_copy)
                inserted_count += 1
            else:
                skipped_count += 1

        if inserted_count > 0:
            flash(f"Successfully loaded {inserted_count} sample student records into MongoDB! ({skipped_count} already existed)", "success")
        else:
            flash("All sample records are already present in the database.", "info")

        return redirect(url_for("index"))

    except PyMongoError as pe:
        flash(f"Failed to load demo data: {str(pe)}", "danger")
        return redirect(url_for("index"))


# =============================================================================
# Application Entry Point
# =============================================================================

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    debug_mode = os.getenv("FLASK_DEBUG", "True").lower() in ("true", "1", "t")

    print("\n" + "=" * 60)
    print(">> Student Record Management System - Flask & MongoDB")
    print("=" * 60)
    print(f"Server running at: http://127.0.0.1:{port}")
    print(f"MongoDB URI:       {MONGO_URI}")
    print(f"Database:          {DB_NAME}")
    print(f"Collection:        {COLLECTION_NAME}")
    print("=" * 60 + "\n")

    app.run(host="0.0.0.0", port=port, debug=debug_mode, use_reloader=False)
