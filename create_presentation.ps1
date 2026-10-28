# ==============================================================================
# Student Record Management System - Professional Presentation Generator
# Generates Student_Record_Management_CA2.pptx and Student_Record_Management_CA2.pdf
# College CA2 Project - Nirmal Kumar
# ==============================================================================

function RGB($r, $g, $b) { 
    return [int]($r + ($g * 256) + ($b * 65536)) 
}

# Color Palette (Academic Blue and White Theme)
$NAVY_DARK     = RGB 15 23 42      # Deep Slate Navy
$NAVY_MED      = RGB 26 54 93      # Academic Navy
$BLUE_PRIMARY  = RGB 37 99 235     # Royal Academic Blue
$BLUE_LIGHT    = RGB 239 246 255   # Soft Ice Blue Card
$SLATE_BG      = RGB 248 250 252   # Card Background
$BORDER_COLOR  = RGB 203 213 225   # Crisp Card Border
$BORDER_ACCENT = RGB 147 197 253   # Soft Blue Border
$TEXT_DARK     = RGB 30 41 59      # Charcoal Main Text
$TEXT_MUTED    = RGB 100 116 139   # Slate Muted Text
$WHITE         = RGB 255 255 255   # Pure White
$GREEN         = RGB 16 185 129    # Emerald Success
$GREEN_BG      = RGB 236 253 245   # Green Tint Card

function Add-Header {
    param($slide, [string]$category, [string]$title)

    # Top Accent Bar
    $bar = $slide.Shapes.AddShape(1, 0, 0, 960, 6)
    $bar.Fill.Solid()
    $bar.Fill.ForeColor.RGB = $BLUE_PRIMARY
    $bar.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse

    # Category Tracker
    $catBox = $slide.Shapes.AddTextbox(1, 45, 20, 870, 20)
    $catBox.TextFrame.MarginLeft = 0
    $catBox.TextFrame.MarginTop = 0
    $catT = $catBox.TextFrame.TextRange
    $catT.Text = $category.ToUpper()
    $catT.Font.Name = "Segoe UI"
    $catT.Font.Size = 10
    $catT.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $catT.Font.Color.RGB = $BLUE_PRIMARY

    # Main Title
    $titleBox = $slide.Shapes.AddTextbox(1, 45, 40, 870, 42)
    $titleBox.TextFrame.MarginLeft = 0
    $titleBox.TextFrame.MarginTop = 0
    $titT = $titleBox.TextFrame.TextRange
    $titT.Text = $title
    $titT.Font.Name = "Segoe UI"
    $titT.Font.Size = 22
    $titT.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $titT.Font.Color.RGB = $NAVY_MED

    # Title Divider Line
    $line = $slide.Shapes.AddShape(1, 45, 88, 870, 1.5)
    $line.Fill.Solid()
    $line.Fill.ForeColor.RGB = $BORDER_COLOR
    $line.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse

    # Footer
    $ftrBox = $slide.Shapes.AddTextbox(1, 45, 510, 870, 20)
    $ftrBox.TextFrame.MarginLeft = 0
    $ftrBox.TextFrame.MarginTop = 0
    $ftrT = $ftrBox.TextFrame.TextRange
    $ftrT.Text = "Student Record Management System  |  College CA2 Project  |  Nirmal Kumar (12201458)"
    $ftrT.Font.Name = "Segoe UI"
    $ftrT.Font.Size = 9
    $ftrT.Font.Color.RGB = $TEXT_MUTED
}

function Add-CardBox {
    param(
        $slide,
        [double]$left,
        [double]$top,
        [double]$width,
        [double]$height,
        [string]$title,
        [string[]]$lines,
        [int]$bgRGB = $SLATE_BG,
        [int]$borderRGB = $BORDER_COLOR,
        [bool]$accent = $false
    )

    $box = $slide.Shapes.AddShape(1, $left, $top, $width, $height)
    $box.Fill.Solid()
    $box.Fill.ForeColor.RGB = $bgRGB
    $box.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $box.Line.ForeColor.RGB = $borderRGB
    $box.Line.Weight = 1.2

    if ($accent) {
        $acc = $slide.Shapes.AddShape(1, $left, $top, 5, $height)
        $acc.Fill.Solid()
        $acc.Fill.ForeColor.RGB = $BLUE_PRIMARY
        $acc.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    }

    $tBox = $slide.Shapes.AddTextbox(1, ($left + 16), ($top + 14), ($width - 32), 26)
    $tBox.TextFrame.MarginLeft = 0
    $tBox.TextFrame.MarginTop = 0
    $tBox.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $tRange = $tBox.TextFrame.TextRange
    $tRange.Text = $title
    $tRange.Font.Name = "Segoe UI"
    $tRange.Font.Size = 13.5
    $tRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $tRange.Font.Color.RGB = $NAVY_DARK

    $bBox = $slide.Shapes.AddTextbox(1, ($left + 16), ($top + 44), ($width - 32), ($height - 52))
    $bBox.TextFrame.MarginLeft = 0
    $bBox.TextFrame.MarginTop = 0
    $bBox.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
    
    $fullText = [string]::Join("`n`n", $lines)
    $bRange = $bBox.TextFrame.TextRange
    $bRange.Text = $fullText
    $bRange.Font.Name = "Segoe UI"
    $bRange.Font.Size = 11
    $bRange.Font.Color.RGB = $TEXT_DARK
}

try {
    Write-Host "Initializing Microsoft PowerPoint COM automation..."
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoFalse)
    $pres.PageSetup.SlideWidth = 960
    $pres.PageSetup.SlideHeight = 540

    $workspaceDir = Get-Location

    # ==========================================================================
    # SLIDE 1: TITLE SLIDE
    # ==========================================================================
    Write-Host "Creating Slide 1: Title Slide..."
    $s1 = $pres.Slides.Add(1, 12)

    # Accent Left Stripe
    $s1Stripe = $s1.Shapes.AddShape(1, 0, 0, 16, 540)
    $s1Stripe.Fill.Solid()
    $s1Stripe.Fill.ForeColor.RGB = $BLUE_PRIMARY
    $s1Stripe.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse

    # Top Pill
    $pill = $s1.Shapes.AddShape(1, 55, 38, 330, 26)
    $pill.Fill.Solid()
    $pill.Fill.ForeColor.RGB = $BLUE_LIGHT
    $pill.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $pill.Line.ForeColor.RGB = $BORDER_ACCENT
    $pt = $pill.TextFrame.TextRange
    $pt.Text = "COLLEGE CA2 PROJECT EVALUATION - 2026"
    $pt.Font.Name = "Segoe UI"
    $pt.Font.Size = 10
    $pt.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $pt.Font.Color.RGB = $BLUE_PRIMARY
    $pt.ParagraphFormat.Alignment = 2

    # Main Project Title
    $s1Title = $s1.Shapes.AddTextbox(1, 55, 76, 850, 70)
    $s1Title.TextFrame.MarginLeft = 0
    $s1T = $s1Title.TextFrame.TextRange
    $s1T.Text = "Student Record Management System"
    $s1T.Font.Name = "Segoe UI"
    $s1T.Font.Size = 33
    $s1T.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $s1T.Font.Color.RGB = $NAVY_DARK

    # Subtitle
    $s1Sub = $s1.Shapes.AddTextbox(1, 55, 150, 850, 32)
    $s1Sub.TextFrame.MarginLeft = 0
    $s1St = $s1Sub.TextFrame.TextRange
    $s1St.Text = "A Robust Full-Stack Web Application Powered by Python, Flask, PyMongo, and MongoDB"
    $s1St.Font.Name = "Segoe UI"
    $s1St.Font.Size = 14.5
    $s1St.Font.Color.RGB = $TEXT_MUTED

    # Student Credentials Card
    $s1Card = $s1.Shapes.AddShape(1, 55, 205, 470, 230)
    $s1Card.Fill.Solid()
    $s1Card.Fill.ForeColor.RGB = $SLATE_BG
    $s1Card.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $s1Card.Line.ForeColor.RGB = $BORDER_COLOR
    $s1Card.Line.Weight = 1.2

    $s1Acc = $s1.Shapes.AddShape(1, 55, 205, 5, 230)
    $s1Acc.Fill.Solid()
    $s1Acc.Fill.ForeColor.RGB = $BLUE_PRIMARY
    $s1Acc.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse

    $s1Head = $s1.Shapes.AddTextbox(1, 75, 218, 430, 25)
    $s1Head.TextFrame.MarginLeft = 0
    $s1H = $s1Head.TextFrame.TextRange
    $s1H.Text = "CANDIDATE AND SUBMISSION CREDENTIALS"
    $s1H.Font.Name = "Segoe UI"
    $s1H.Font.Size = 11
    $s1H.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $s1H.Font.Color.RGB = $BLUE_PRIMARY

    $s1Body = $s1.Shapes.AddTextbox(1, 75, 248, 430, 175)
    $s1Body.TextFrame.MarginLeft = 0
    $s1B = $s1Body.TextFrame.TextRange
    $s1B.Text = "Candidate Name:  Nirmal Kumar`nRegistration No:  12201458`nCourse Program:  B.Tech (Computer Science and Engineering)`nDepartment:      Department of Computer Science and Engineering`nEvaluation Mode: Continuous Assessment 2 (CA2)`nAcademic Term:   Spring 2026"
    $s1B.Font.Name = "Segoe UI"
    $s1B.Font.Size = 12
    $s1B.Font.Color.RGB = $TEXT_DARK

    # Tech Stack Card
    $s1Tech = $s1.Shapes.AddShape(1, 545, 205, 360, 230)
    $s1Tech.Fill.Solid()
    $s1Tech.Fill.ForeColor.RGB = $BLUE_LIGHT
    $s1Tech.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $s1Tech.Line.ForeColor.RGB = $BORDER_ACCENT
    $s1Tech.Line.Weight = 1.2

    $s1Th = $s1.Shapes.AddTextbox(1, 565, 218, 320, 25)
    $s1Th.TextFrame.MarginLeft = 0
    $s1ThT = $s1Th.TextFrame.TextRange
    $s1ThT.Text = "CORE TECHNOLOGIES AND TOOLS"
    $s1ThT.Font.Name = "Segoe UI"
    $s1ThT.Font.Size = 11
    $s1ThT.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $s1ThT.Font.Color.RGB = $NAVY_MED

    $s1Tb = $s1.Shapes.AddTextbox(1, 565, 248, 320, 175)
    $s1Tb.TextFrame.MarginLeft = 0
    $s1TbT = $s1Tb.TextFrame.TextRange
    $s1TbT.Text = "[x] Python 3.10  (Application Backend Logic)`n[x] Flask 3.1    (WSGI Micro Web Framework)`n[x] MongoDB      (NoSQL Document Database)`n[x] PyMongo 4.18 (Python Database Driver)`n[x] HTML5 and CSS3 (Responsive UI and Cards)`n[x] Python-Dotenv (Secure Environment Config)"
    $s1TbT.Font.Name = "Segoe UI"
    $s1TbT.Font.Size = 11.5
    $s1TbT.Font.Color.RGB = $NAVY_DARK

    # ==========================================================================
    # SLIDE 2: INTRODUCTION
    # ==========================================================================
    Write-Host "Creating Slide 2: Introduction..."
    $s2 = $pres.Slides.Add(2, 12)
    Add-Header $s2 "01 | Introduction" "Modernizing Academic Record Management"

    Add-CardBox $s2 45 105 420 385 "Background and Context" @(
        "Academic institutions handle vast quantities of student academic, personal, and enrollment data across departments every term.",
        "Legacy reliance on physical paper registers and unlinked spreadsheets introduces data duplication, calculation mistakes, and slow retrieval.",
        "The Student Record Management System provides a centralized, secure web-based registry accessible via any modern web browser.",
        "Developed specifically for college continuous assessment (CA2), offering complete student lifecycle management with zero operational friction."
    ) $SLATE_BG $BORDER_COLOR $true

    Add-CardBox $s2 495 105 420 385 "Key Advantages and Architecture" @(
        "Zero Installation Barrier: Runs smoothly in any web browser without requiring heavy desktop client software.",
        "Document-Oriented Storage: Utilizes MongoDB BSON documents instead of rigid tables, reflecting natural student profile structures.",
        "High Fault Tolerance: Built-in connection ping check, server fault isolation, and safe input parsing prevent application crashes.",
        "Live KPI Analytics: Instant visual calculation of total enrolled students, overall average CGPA, and top academic scores upon every visit."
    ) $BLUE_LIGHT $BORDER_ACCENT $true

    # ==========================================================================
    # SLIDE 3: PROBLEM STATEMENT
    # ==========================================================================
    Write-Host "Creating Slide 3: Problem Statement..."
    $s3 = $pres.Slides.Add(3, 12)
    Add-Header $s3 "02 | Problem Statement" "Challenges in Traditional Student Record Systems"

    Add-CardBox $s3 45 105 275 385 "1. Data Redundancy" @(
        "Duplicate Enrollments: Manual entry frequently leads to duplicate roll numbers assigned to multiple students.",
        "Inconsistent Records: Changes made in local spreadsheets fail to propagate to department heads.",
        "Rigid Schemas: Adding custom fields in legacy SQL systems requires complex, disruptive database migrations."
    ) $SLATE_BG $BORDER_COLOR $true

    Add-CardBox $s3 342 105 275 385 "2. Inefficient Retrieval" @(
        "Slow Querying: Manually searching through hundreds of paper records or spreadsheet rows causes delays.",
        "No Filter Controls: Inability to instantly filter student cohorts by academic department or semester standing.",
        "Fragmented Analytics: Lack of high-level overview metrics such as class CGPA averages and department enrollment numbers."
    ) $SLATE_BG $BORDER_COLOR $true

    Add-CardBox $s3 640 105 275 385 "3. Data Safety and Integrity" @(
        "Accidental Deletion: Unprotected record modification with no confirmation prompts or undo awareness.",
        "Validation Omissions: Invalid CGPAs (e.g. greater than 10.0) or malformed emails enter the system unchecked.",
        "Hardcoded Secrets: Storing database passwords directly in source code creates severe security vulnerabilities."
    ) $SLATE_BG $BORDER_COLOR $true

    # ==========================================================================
    # SLIDE 4: PROJECT OBJECTIVES
    # ==========================================================================
    Write-Host "Creating Slide 4: Project Objectives..."
    $s4 = $pres.Slides.Add(4, 12)
    Add-Header $s4 "03 | Project Objectives" "Key Goals and Architectural Milestones"

    Add-CardBox $s4 45 105 420 180 "1. Complete CRUD Functionality" @(
        "Deliver complete Create, Read, Update, and Delete operations for student academic profiles.",
        "Ensure instantaneous user interface updates with automated HTTP redirects and flash alerts."
    ) $SLATE_BG $BORDER_COLOR $true

    Add-CardBox $s4 495 105 420 180 "2. Strict Data Integrity and Validation" @(
        "Enforce unique registration numbers at both application logic and MongoDB index layers.",
        "Validate CGPA bounds (0.00 to 10.00), semester limits (1 to 8), and institutional email syntax."
    ) $SLATE_BG $BORDER_COLOR $true

    Add-CardBox $s4 45 305 420 185 "3. Real-Time Search and Dashboard" @(
        "Provide regex-powered live searching by registration number or student name.",
        "Display dynamic analytics cards: Total Students, Average CGPA, Highest CGPA, and Active Courses."
    ) $BLUE_LIGHT $BORDER_ACCENT $true

    Add-CardBox $s4 495 305 420 185 "4. Production-Grade Configuration" @(
        "Isolate environment settings via .env and provide sample templates (.env.example).",
        "Guarantee database fault tolerance with automatic connection recovery and ping monitoring."
    ) $BLUE_LIGHT $BORDER_ACCENT $true

    # ==========================================================================
    # SLIDE 5: TECHNOLOGIES USED AND THEIR ROLES
    # ==========================================================================
    Write-Host "Creating Slide 5: Technologies Used..."
    $s5 = $pres.Slides.Add(5, 12)
    Add-Header $s5 "04 | Technology Stack" "Technologies Used and Their Architectural Roles"

    Add-CardBox $s5 45 105 420 180 "Python 3.10  (Core Backend)" @(
        "Core programming language powering backend business logic and input sanitation.",
        "Handles mathematical rounding, regex validation, and environment variable parsing."
    ) $SLATE_BG $BORDER_COLOR $true

    Add-CardBox $s5 495 105 420 180 "Flask 3.1  (WSGI Web Framework)" @(
        "Lightweight WSGI framework handling RESTful routing (@app.route decorators).",
        "Manages Jinja2 HTML rendering, session cookies, and contextual flash alert notifications."
    ) $SLATE_BG $BORDER_COLOR $true

    Add-CardBox $s5 45 305 420 185 "MongoDB and PyMongo 4.18  (Database Layer)" @(
        "High-throughput NoSQL document database storing BSON student records.",
        "PyMongo driver provides connection pooling, CRUD queries, and unique index enforcement."
    ) $SLATE_BG $BORDER_COLOR $true

    Add-CardBox $s5 495 305 420 185 "HTML5 and Vanilla CSS3  (Frontend Layer)" @(
        "Responsive, clean layout built with CSS custom properties and Google Fonts.",
        "Features interactive data tables, CGPA performance badges, and confirmation modals."
    ) $SLATE_BG $BORDER_COLOR $true

    # ==========================================================================
    # SLIDE 6: SYSTEM ARCHITECTURE
    # ==========================================================================
    Write-Host "Creating Slide 6: System Architecture..."
    $s6 = $pres.Slides.Add(6, 12)
    Add-Header $s6 "05 | System Architecture" "Multi-Tier Web Application Workflow"

    Add-CardBox $s6 45 115 200 375 "1. Client Layer`n(Browser / UI)" @(
        "HTML5 Templates:`n(index, add, edit)",
        "Responsive CSS3 Styling",
        "User HTTP Requests`n(GET and POST)",
        "Form Submissions",
        "JS Confirmation Alerts",
        "Flash Notification Banners"
    ) $BLUE_LIGHT $BORDER_ACCENT $false

    $ar1 = $s6.Shapes.AddShape(33, 253, 280, 32, 22)
    $ar1.Fill.Solid()
    $ar1.Fill.ForeColor.RGB = $BLUE_PRIMARY
    $ar1.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse

    Add-CardBox $s6 295 115 200 375 "2. Server Layer`n(Flask / Python)" @(
        "app.py Application",
        "URL Routing Engine`n(/, /add, /edit, /delete)",
        "Regex Data Sanitizer",
        "Duplicate Validator",
        "Statistics Calculator",
        "Jinja2 Template Engine"
    ) $SLATE_BG $BORDER_COLOR $false

    $ar2 = $s6.Shapes.AddShape(33, 503, 280, 32, 22)
    $ar2.Fill.Solid()
    $ar2.Fill.ForeColor.RGB = $BLUE_PRIMARY
    $ar2.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse

    Add-CardBox $s6 545 115 170 375 "3. Driver Layer`n(PyMongo)" @(
        "MongoClient Pool",
        "BSON Object Mapping",
        "ObjectId Translation",
        "Query Execution:`n(insert_one, find,`n update_one, delete)",
        "Index Management",
        "Error Handling"
    ) $SLATE_BG $BORDER_COLOR $false

    $ar3 = $s6.Shapes.AddShape(33, 723, 280, 32, 22)
    $ar3.Fill.Solid()
    $ar3.Fill.ForeColor.RGB = $BLUE_PRIMARY
    $ar3.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse

    Add-CardBox $s6 765 115 150 375 "4. Data Store`n(MongoDB)" @(
        "Database:`nstudent_records_db",
        "Collection:`nstudents",
        "Document Storage",
        "Unique Index`non reg_no",
        "WiredTiger Engine",
        "Persistent Disk"
    ) $GREEN_BG (RGB 167 243 208) $false

    # ==========================================================================
    # SLIDE 7: MAIN FEATURES
    # ==========================================================================
    Write-Host "Creating Slide 7: Main Features..."
    $s7 = $pres.Slides.Add(7, 12)
    Add-Header $s7 "06 | Main Features" "Core CRUD Operations and Application Capabilities"

    Add-CardBox $s7 45 105 275 180 "Add Student (CREATE)" @(
        "Route: POST /add",
        "Captures: Name, Reg No, Course, Semester, Email, and CGPA.",
        "Uniqueness Check: Queries MongoDB to block duplicate reg_no."
    ) $SLATE_BG $BORDER_COLOR $true

    Add-CardBox $s7 342 105 275 180 "View Registry (READ)" @(
        "Route: GET /",
        "Formatted table with custom badges for courses and semesters.",
        "Color-coded CGPA pills: Outstanding (>=8.5), Good (>=7.0), Low."
    ) $SLATE_BG $BORDER_COLOR $true

    Add-CardBox $s7 640 105 275 180 "Search and Filter (QUERY)" @(
        "Routes: GET / and GET /search",
        "Case-insensitive regex search by Reg Number or Name.",
        "Course filter dropdown plus instant Clear Filters reset button."
    ) $SLATE_BG $BORDER_COLOR $true

    Add-CardBox $s7 185 305 285 185 "Update Record (UPDATE)" @(
        "Route: POST /edit/<id>",
        "Pre-fills existing student credentials into edit template.",
        "Conflict Prevention: Ensures updated reg_no is not held by another student."
    ) $BLUE_LIGHT $BORDER_ACCENT $true

    Add-CardBox $s7 490 305 285 185 "Delete Record (DELETE)" @(
        "Route: POST /delete/<id>",
        "Client-side safety confirmation: JavaScript onsubmit confirm() prompt.",
        "Permanent document removal followed by friendly flash notice."
    ) $BLUE_LIGHT $BORDER_ACCENT $true

    # ==========================================================================
    # SLIDE 8: DATABASE DESIGN
    # ==========================================================================
    Write-Host "Creating Slide 8: Database Design..."
    $s8 = $pres.Slides.Add(8, 12)
    Add-Header $s8 "07 | Database Design" "MongoDB Collection and BSON Document Schema"

    Add-CardBox $s8 45 105 420 385 "Database and Collection Architecture" @(
        "Database Name: student_records_db",
        "Collection Name: students",
        "Storage Format: BSON (Binary JSON) documents",
        "Index Rule: create_index([('reg_no', 1)], unique=True)",
        "Schema Flexibility: Allows adding timestamps (created_at, updated_at) without schema alteration scripts.",
        "High Performance: WiredTiger storage engine handles concurrent document reads and writes efficiently without table lockups."
    ) $SLATE_BG $BORDER_COLOR $true

    Add-CardBox $s8 495 105 420 385 "Document Schema and Field Types" @(
        "_id : ObjectId (12-byte primary key generated by MongoDB)",
        "name : String (Full legal name of enrolled student)",
        "reg_no : String (Unique student registration number)",
        "course : String (Academic program, e.g. B.Tech CSE, BCA)",
        "semester : Integer (Current term, integer between 1 and 8)",
        "email : String (Valid institutional email address)",
        "cgpa : Double (Cumulative GPA on a scale of 0.00 to 10.00)",
        "created_at : DateTime (Timestamp of initial document creation)"
    ) $BLUE_LIGHT $BORDER_ACCENT $true

    # ==========================================================================
    # SLIDE 9: PROJECT SCREENSHOTS
    # ==========================================================================
    Write-Host "Creating Slide 9: Project Screenshots..."
    $s9 = $pres.Slides.Add(9, 12)
    Add-Header $s9 "08 | Project Screenshots and Output" "Live Application Demonstration and Actual Output Screenshots"

    $dashImgPath = Join-Path $workspaceDir "screenshots\dashboard.png"
    $searchImgPath = Join-Path $workspaceDir "screenshots\search_results.png"

    if (Test-Path $dashImgPath) {
        $f1 = $s9.Shapes.AddShape(1, 45, 105, 420, 255)
        $f1.Fill.Solid()
        $f1.Fill.ForeColor.RGB = $WHITE
        $f1.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $f1.Line.ForeColor.RGB = $BORDER_COLOR
        $f1.Line.Weight = 1.2

        $pic1 = $s9.Shapes.AddPicture($dashImgPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 47, 107, 416, 251)

        $c1Box = $s9.Shapes.AddShape(1, 45, 370, 420, 115)
        $c1Box.Fill.Solid()
        $c1Box.Fill.ForeColor.RGB = $SLATE_BG
        $c1Box.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $c1Box.Line.ForeColor.RGB = $BORDER_COLOR
        $c1T = $c1Box.TextFrame.TextRange
        $c1T.Text = "Figure 1: Main Dashboard and Student Registry`n- Live Analytics: Total Students, Avg CGPA, Highest CGPA, Courses`n- MongoDB Connected badge verifying active database status`n- Responsive registry table with action buttons and CGPA pills"
        $c1T.Font.Name = "Segoe UI"
        $c1T.Font.Size = 10
        $c1T.Font.Color.RGB = $TEXT_DARK
    }

    if (Test-Path $searchImgPath) {
        $f2 = $s9.Shapes.AddShape(1, 495, 105, 420, 255)
        $f2.Fill.Solid()
        $f2.Fill.ForeColor.RGB = $WHITE
        $f2.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $f2.Line.ForeColor.RGB = $BORDER_COLOR
        $f2.Line.Weight = 1.2

        $pic2 = $s9.Shapes.AddPicture($searchImgPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 497, 107, 416, 251)

        $c2Box = $s9.Shapes.AddShape(1, 495, 370, 420, 115)
        $c2Box.Fill.Solid()
        $c2Box.Fill.ForeColor.RGB = $BLUE_LIGHT
        $c2Box.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $c2Box.Line.ForeColor.RGB = $BORDER_ACCENT
        $c2T = $c2Box.TextFrame.TextRange
        $c2T.Text = "Figure 2: Real-Time Search and Query Operation`n- Instant lookup matching Registration Number (2024BCSE001)`n- Dynamic record counter badge ('Showing 1 of 6 Records')`n- Instant 'Clear Filters' button returning to the full registry"
        $c2T.Font.Name = "Segoe UI"
        $c2T.Font.Size = 10
        $c2T.Font.Color.RGB = $NAVY_DARK
    }

    # ==========================================================================
    # SLIDE 10: CONCLUSION AND FUTURE SCOPE
    # ==========================================================================
    Write-Host "Creating Slide 10: Conclusion and Future Scope..."
    $s10 = $pres.Slides.Add(10, 12)
    Add-Header $s10 "09 | Conclusion and Future Scope" "Project Summary and Future Enhancement Roadmap"

    Add-CardBox $s10 45 105 420 385 "Project Summary and Achievements" @(
        "[x] Complete CRUD Implementation: Fully functional web application satisfying all college CA2 requirements.",
        "[x] 100% Automated Test Pass: All 7 integration tests verified for route availability, duplicate prevention, and bounds.",
        "[x] Robust Database Architecture: MongoDB NoSQL backend with automatic reconnection, index uniqueness, and schema resilience.",
        "[x] User-Centric Design: Clean academic dashboard featuring live analytics, color-coded grade badges, and intuitive navigation."
    ) $GREEN_BG (RGB 167 243 208) $true

    Add-CardBox $s10 495 105 420 385 "Future Enhancement Roadmap" @(
        "Role-Based Access Control (RBAC): Dedicated portals and permissions for Administrators, Faculty, and Students.",
        "Automated Marksheet Generation: Exporting semester report cards into downloadable PDF certificates.",
        "Attendance and Fee Tracking: Extending the schema to record class attendance percentages and college tuition clearance.",
        "REST API and Mobile Integration: Exposing JSON endpoints for future cross-platform Android and iOS mobile apps."
    ) $BLUE_LIGHT $BORDER_ACCENT $true

    # ==========================================================================
    # SAVE PPTX AND PDF
    # ==========================================================================
    $pptxOutput = Join-Path $workspaceDir "Student_Record_Management_CA2.pptx"
    $pdfOutput  = Join-Path $workspaceDir "Student_Record_Management_CA2.pdf"

    Write-Host "Saving presentation to PPTX: $pptxOutput ..."
    $pres.SaveAs($pptxOutput)

    Write-Host "Exporting presentation to PDF: $pdfOutput ..."
    $pres.SaveAs($pdfOutput, 32) # 32 = ppSaveAsPDF

    $pres.Close()
    $ppt.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null

    Write-Host "`n========================================================"
    Write-Host "Presentation Generation Successful!"
    Write-Host "PPTX Path: $pptxOutput"
    Write-Host "PDF  Path: $pdfOutput"
    Write-Host "Total Slides: 10"
    Write-Host "========================================================`n"

} catch {
    Write-Host "Generation Error: $_"
    if ($ppt) { $ppt.Quit() }
}
