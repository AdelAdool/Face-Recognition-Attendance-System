# FaceTrack — AI Attendance System

## Requirements
- Python 3.8+
- Webcam
- Windows / Mac / Linux

## Quick Setup

### Step 1 — Install dependencies
```
pip install -r requirements.txt
```

> Note: `face-recognition` requires `dlib`. On Windows you may need to install it separately:
> ```
> pip install cmake
> pip install dlib
> pip install face-recognition
> ```
> Or use the prebuilt wheel: https://github.com/z-mahmud22/Dlib_Windows_Python3.x

### Step 2 — Run the server
```
python app.py
```

### Step 3 — Open the app
Go to: **http://localhost:5000**

---

## How to Use

### Register a Person
1. Click the **Register** tab
2. Click **Start Camera**
3. Type the person's name
4. Click **📸 Capture** (face clearly in frame)
5. Click **✓ Register This Person**

### Mark Attendance
1. Click the **Attendance** tab
2. Click **Start Camera**
3. Either:
   - Click **Scan Now** to check once
   - Toggle **Auto-scan** to check every 3 seconds automatically
4. Names and timestamps are automatically logged

### Export to Excel
- Go to the **Log** tab
- Click **⬇ Export Excel**
- An `.xlsx` file downloads with two sheets:
  - **Attendance** — full log with name, date, time, status
  - **Summary** — totals per person

---

## Files Created
- `face_encodings.pkl` — stored face data (auto-created)
- `attendance_YYYYMMDD_HHMMSS.xlsx` — exported attendance files
