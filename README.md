# 🪑 Seating and Attendance Sheet Generator

A Python GUI tool to **generate seating arrangements** and **attendance sheets** for exam rooms using a layout file and student roll number lists.

![Main UI](./assets/Screenshot_Main.png)

---

## 🎯 Features

- Upload room layout Excel file.
- Upload roll number files for left, middle, and right blocks.
- Generate:
  - Seating charts per room.
  - Attendance sheets per room.
- Export results to an Excel file with multiple sheets.
- Simple graphical interface using **Tkinter**.

---

## 📥 Input Requirements

- **Room Layout File**: Defines the rows, columns, and structure of rooms.
- **Roll Number Files**:
  - `Left` side student roll numbers.
  - `Middle` side student roll numbers.
  - `Right` side student roll numbers.

> All files must be in **Excel (.xlsx)** format.
---

## 📤 Output

An Excel file (`SeatingChart_Output.xlsx`) containing:
- Individual seating arrangement sheets (e.g., `Room 302`)
- Attendance sheets for each room (e.g., `Attendance - Room 302`)

![Excel Output](./assets/Screenshot_Output.png)

---

## 📽 Demo Video

> Watch the tool in action 👇  
[![Watch Video](./assets/video_thumbnail.png)](https://github.com/Goutam16-Withcode/Arrangement_Chart/assets/demo.mp4)

---

## 🖼 Interface Screenshots

| Step | Screenshot |
|------|------------|
| GUI Interface | ![Screenshot_Main](./assets/Screenshot_Main.png) |
| File Selection | ![File Selection](./assets/Screenshot_File_Select.png) |
| Output in Excel | ![Output Excel](./assets/Screenshot_Output.png) |

---

## 🚀 How to Run

1. **Clone this repository**
   ```bash
   git clone https://github.com/Goutam16-Withcode/Arrangement_Chart
   cd Arrangement_Chart
