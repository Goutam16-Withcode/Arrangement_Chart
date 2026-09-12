import * as XLSX from "xlsx";
import { RoomConfig, Student, RoomSeating } from "./types";

export class ExcelEngine {
  /**
   * Export seating charts and attendance sheets to a multi-sheet .xlsx workbook
   */
  public static exportSeatingWorkbook(roomSeatings: RoomSeating[]): void {
    const wb = XLSX.utils.book_new();

    roomSeatings.forEach((rs) => {
      const room = rs.roomConfig;
      const roomTitle = `Room ${room.roomNumber}`;

      // 1. Build Seating Chart Sheet Data
      const seatingData: any[][] = [];

      // Header Banner
      seatingData.push([`EXAMINATION SEATING ARRANGEMENT - ROOM ${room.roomNumber}`]);
      seatingData.push([`Building: ${room.building} | Floor: ${room.floor} | Total Capacity: ${rs.totalCapacity}`]);
      seatingData.push([]);

      // Seat Position Headers (e.g. Bench 1, Bench 2, ...)
      const benchHeaders: string[] = [];
      for (let b = 0; b < room.benchesPerRow; b++) {
        for (let p = 0; p < room.studentsPerBench; p++) {
          const pos = ["F-1", "S-1", "T-1", "F-2"][p] || `P-${p + 1}`;
          benchHeaders.push(`Col ${b + 1} (${pos})`);
        }
      }
      seatingData.push(benchHeaders);

      // Rows of benches
      for (let r = 0; r < room.rows; r++) {
        const rowData: string[] = [];
        for (let c = 0; c < room.benchesPerRow; c++) {
          const benchIndex = r * room.benchesPerRow + c;
          const bench = rs.seats[benchIndex];

          for (let p = 0; p < room.studentsPerBench; p++) {
            if (bench && bench[p] && bench[p].student) {
              rowData.push(`${bench[p].student!.rollNo} (${bench[p].student!.branch})`);
            } else {
              rowData.push("--- [VACANT] ---");
            }
          }
        }
        seatingData.push(rowData);
      }

      const wsSeating = XLSX.utils.aoa_to_sheet(seatingData);
      XLSX.utils.book_append_sheet(wb, wsSeating, roomTitle);

      // 2. Build Attendance Sheet
      const attendanceData: any[][] = [
        [`ATTENDANCE SHEET - ROOM ${room.roomNumber}`],
        [`Course Exam Session | Total Students Seated: ${rs.assignedCount}`],
        [],
        ["Serial No", "Seat Position", "Roll Number", "Student Name", "Branch", "Signature / Status"],
      ];

      rs.attendanceList.forEach((item) => {
        attendanceData.push([
          item.serialNo,
          item.position,
          item.student.rollNo,
          item.student.name,
          item.student.branch,
          item.present ? "PRESENT" : "____________________",
        ]);
      });

      const wsAttendance = XLSX.utils.aoa_to_sheet(attendanceData);
      XLSX.utils.book_append_sheet(wb, wsAttendance, `Att - Room ${room.roomNumber}`);
    });

    // Write file and trigger browser download
    XLSX.writeFile(wb, "SeatingChart_Smart_Output.xlsx");
  }

  /**
   * Parse uploaded Room Layout Excel file
   */
  public static async parseRoomLayoutFile(file: File): Promise<RoomConfig[]> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: "array" });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const json: any[] = XLSX.utils.sheet_to_json(firstSheet);

          const rooms: RoomConfig[] = json.map((row, idx) => ({
            roomNumber: String(row["Room Number"] || row["room_number"] || `Room-${idx + 1}`),
            building: String(row["Building"] || "Main Academic Complex"),
            floor: Number(row["Floor"] || 1),
            rows: Number(row["Number of Rows"] || row["rows"] || 5),
            benchesPerRow: Number(row["Number of Bench"] || row["benches"] || 4),
            studentsPerBench: Number(row["Number of Student per Bench"] || row["students_per_bench"] || 3),
            leftBranchName: row["Left Name"] || "Branch 1",
            middleBranchName: row["Middle Name"] || "Branch 2",
            rightBranchName: row["Right Name"] || "Branch 3",
          }));

          resolve(rooms);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  /**
   * Parse uploaded Roll Numbers Excel file
   */
  public static async parseRollNumbersFile(file: File, branchName: string): Promise<Student[]> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: "array" });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const json: any[] = XLSX.utils.sheet_to_json(firstSheet);

          const students: Student[] = json.map((row, idx) => ({
            id: `upload-${branchName}-${idx + 1}`,
            rollNo: String(row["Roll Number"] || row["roll_number"] || row["RollNo"] || `ROLL-${idx + 1}`),
            name: String(row["Name"] || row["Student Name"] || `Student ${idx + 1}`),
            branch: branchName,
            year: Number(row["Year"] || 2),
            subjectCode: String(row["Subject Code"] || `${branchName}101`),
            subjectName: String(row["Subject"] || "Core Subject"),
          }));

          resolve(students);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }
}
