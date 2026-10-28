import os
from datetime import datetime
from dotenv import load_dotenv
import pymongo

load_dotenv()

MONGO_URI = os.getenv("MONGO_URI", "mongodb://localhost:27017/")
DB_NAME = os.getenv("DB_NAME", "student_records_db")
COLLECTION_NAME = os.getenv("COLLECTION_NAME", "students")

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
    },
    {
        "name": "Ananya Roy",
        "reg_no": "2024BCSE088",
        "course": "B.Tech CSE",
        "semester": 6,
        "email": "ananya.roy@college.edu",
        "cgpa": 9.05,
        "created_at": datetime.now()
    }
]


def seed_database():
    print(f"Connecting to MongoDB at: {MONGO_URI} ...")
    try:
        client = pymongo.MongoClient(MONGO_URI, serverSelectionTimeoutMS=3000)
        client.admin.command('ping')
        db = client[DB_NAME]
        collection = db[COLLECTION_NAME]

        # Ensure unique index
        collection.create_index([("reg_no", pymongo.ASCENDING)], unique=True)

        print(f"Connected to database '{DB_NAME}', collection '{COLLECTION_NAME}'.")

        inserted = 0
        skipped = 0

        for student in DEMO_STUDENTS:
            existing = collection.find_one({"reg_no": student["reg_no"]})
            if existing:
                print(f"  [SKIPPED] {student['name']} ({student['reg_no']}) - already exists.")
                skipped += 1
            else:
                collection.insert_one(student)
                print(f"  [INSERTED] {student['name']} ({student['reg_no']}) - CGPA: {student['cgpa']}")
                inserted += 1

        print("\n" + "=" * 50)
        print(f"Seed Summary: {inserted} added, {skipped} skipped.")
        print(f"Total students now in database: {collection.count_documents({})}")
        print("=" * 50)

    except Exception as e:
        print(f"Error seeding database: {e}")


if __name__ == "__main__":
    seed_database()
