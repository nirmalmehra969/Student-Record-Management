import pytest
from app import app, get_db
from bson.objectid import ObjectId


@pytest.fixture
def client():
    app.config["TESTING"] = True
    app.config["WTF_CSRF_ENABLED"] = False
    with app.test_client() as client:
        yield client


def test_index_page(client):
    """Test Home Page (Dashboard and Student Table)"""
    response = client.get("/")
    assert response.status_code == 200
    assert b"Student Record Management" in response.data
    assert b"Total Students" in response.data
    assert b"Average CGPA" in response.data


def test_search_student_by_reg_no(client):
    """Test Search by Registration Number"""
    response = client.get("/?search=2024BCSE001")
    assert response.status_code == 200
    assert b"Aarav Sharma" in response.data
    assert b"2024BCSE001" in response.data


def test_search_student_by_name(client):
    """Test Search by Student Name"""
    response = client.get("/?search=Priya")
    assert response.status_code == 200
    assert b"Priya Patel" in response.data


def test_add_student_get(client):
    """Test GET /add page renders form"""
    response = client.get("/add")
    assert response.status_code == 200
    assert b"Add New Student Record" in response.data
    assert b"Registration Number" in response.data


def test_add_student_duplicate_prevention(client):
    """Test adding duplicate registration number fails with warning"""
    duplicate_payload = {
        "name": "Clone Student",
        "reg_no": "2024BCSE001",  # Already exists
        "course": "B.Tech CSE",
        "semester": 4,
        "email": "clone@college.edu",
        "cgpa": "8.50"
    }
    response = client.post("/add", data=duplicate_payload, follow_redirects=True)
    assert response.status_code == 200
    assert b"already exists" in response.data


def test_add_student_invalid_cgpa(client):
    """Test invalid CGPA (>10.0) is rejected"""
    invalid_payload = {
        "name": "Test Student",
        "reg_no": "2024TEST999",
        "course": "BCA",
        "semester": 2,
        "email": "test@college.edu",
        "cgpa": "11.5"  # Invalid!
    }
    response = client.post("/add", data=invalid_payload, follow_redirects=True)
    assert response.status_code == 200
    assert b"CGPA must be between 0.00 and 10.00" in response.data


def test_crud_lifecycle(client):
    """Test Full CRUD Lifecycle: Create -> Read -> Update -> Delete"""
    collection, _ = get_db()
    test_reg_no = "2024LIFECYCLE01"
    
    # Clean up before test if exists
    collection.delete_one({"reg_no": test_reg_no})

    # 1. CREATE
    payload = {
        "name": "Rohit Lifecycle",
        "reg_no": test_reg_no,
        "course": "BCA",
        "semester": 3,
        "email": "rohit.life@college.edu",
        "cgpa": "8.40"
    }
    post_res = client.post("/add", data=payload, follow_redirects=True)
    assert post_res.status_code == 200
    assert b"Rohit Lifecycle" in post_res.data
    assert b"added successfully" in post_res.data

    # Fetch created student from DB
    created = collection.find_one({"reg_no": test_reg_no})
    assert created is not None
    student_id = str(created["_id"])

    # 2. READ / EDIT PAGE GET
    edit_get = client.get(f"/edit/{student_id}")
    assert edit_get.status_code == 200
    assert b"Rohit Lifecycle" in edit_get.data

    # 3. UPDATE
    update_payload = {
        "name": "Rohit Updated",
        "reg_no": test_reg_no,
        "course": "MCA",
        "semester": 4,
        "email": "rohit.updated@college.edu",
        "cgpa": "9.10"
    }
    update_res = client.post(f"/edit/{student_id}", data=update_payload, follow_redirects=True)
    assert update_res.status_code == 200
    assert b"updated successfully" in update_res.data

    updated = collection.find_one({"_id": ObjectId(student_id)})
    assert updated["name"] == "Rohit Updated"
    assert updated["course"] == "MCA"
    assert updated["cgpa"] == 9.10

    # 4. DELETE
    delete_res = client.post(f"/delete/{student_id}", follow_redirects=True)
    assert delete_res.status_code == 200
    assert b"permanently deleted" in delete_res.data

    # Verify deleted from DB
    deleted = collection.find_one({"_id": ObjectId(student_id)})
    assert deleted is None
