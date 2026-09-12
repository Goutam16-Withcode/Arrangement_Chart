import * as XLSX from "xlsx";
import { RoomConfig, Student, RoomSeating, CollegeProfile, ExamSession } from "./types";

export class ExcelEngine {
  /**
   * Export seating charts and attendance sheets to a multi-sheet .xlsx workbook
   * with Section-Wise grouped attendance records (F-1, S-1, T-1 / Branch sections).
   */
  public static exportSeatingWorkbook(
    roomSeatings: RoomSeating[],
    collegeProfile?: CollegeProfile,
    activeSession?: ExamSession
  ): void {
    const wb = XLSX.utils.book_new();

    const collegeName = collegeProfile?.collegeName || "National Institute of Engineering & Technology";
    const collegeCode = collegeProfile?.collegeCode || "NIET-1084";
    const examTitle = activeSession?.title || "End-Semester Major Theory Exam";
    const examDate = activeSession?.date || "28-Nov-2026";
    const examTiming = activeSession?.timing || "09:30 AM - 12:30 PM";

    roomSeatings.forEach((rs) => {
      const room = rs.roomConfig;
      const roomTitle = `Room ${room.roomNumber}`;

      // 1. Build Seating Chart Sheet Data
      const seatingData: any[][] = [];

      // Header Banner
      seatingData.push([collegeName.toUpperCase()]);
      seatingData.push([`INSTITUTE CODE: ${collegeCode} | EXAM: ${examTitle.toUpperCase()}`]);
      seatingData.push([`DATE: ${examDate} | TIMING: ${examTiming} | ROOM: ${room.roomNumber}`]);
      seatingData.push([`Building: ${room.building} | Floor: ${room.floor} | Total Capacity: ${rs.totalCapacity}`]);
      seatingData.push([]);

      // Seat Position Headers (e.g. Col 1 (F-1), Col 1 (S-1), ...)
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

      // 2. Build Section-Wise Attendance Sheet
      const sectionPositions = ["F-1", "S-1", "T-1", "F-2"].slice(0, room.studentsPerBench);

      const sectionGroups: { [pos: string]: typeof rs.attendanceList } = {};
      sectionPositions.forEach((pos) => {
        sectionGroups[pos] = rs.attendanceList.filter((item) => item.position === pos);
      });

      const maxRows = Math.max(
        ...sectionPositions.map((pos) => sectionGroups[pos]?.length || 0),
        1
      );

      const attendanceData: any[][] = [];
      attendanceData.push([collegeName.toUpperCase()]);
      attendanceData.push([`SECTION-WISE CANDIDATE ATTENDANCE REGISTER • ${examTitle.toUpperCase()}`]);
      attendanceData.push([`DATE: ${examDate} | TIME: ${examTiming} | ROOM ${room.roomNumber} (${rs.assignedCount} Candidates)`]);
      attendanceData.push([]);

      // Section Header Banner Row
      const secHeaderRow: string[] = [];
      const colHeaderRow: string[] = [];

      sectionPositions.forEach((pos) => {
        const posLabel =
          pos === "F-1"
            ? `SECTION 1 (LEFT - ${pos})`
            : pos === "S-1"
            ? `SECTION 2 (MIDDLE - ${pos})`
            : pos === "T-1"
            ? `SECTION 3 (RIGHT - ${pos})`
            : `SECTION (${pos})`;

        secHeaderRow.push(posLabel, "", "", "", "");
        colHeaderRow.push("S.No", "Roll Number", "Student Name", "Branch", "Signature / Status");
      });

      attendanceData.push(secHeaderRow);
      attendanceData.push(colHeaderRow);

      // Rows for side-by-side section tables
      for (let i = 0; i < maxRows; i++) {
        const row: string[] = [];

        sectionPositions.forEach((pos) => {
          const item = sectionGroups[pos]?.[i];
          if (item) {
            row.push(
              String(i + 1),
              item.student.rollNo,
              item.student.name,
              item.student.branch,
              item.present ? "PRESENT" : "________________"
            );
          } else {
            row.push("", "", "", "", "");
          }
        });

        attendanceData.push(row);
      }

      // Add Branch Summary Block at the bottom
      attendanceData.push([]);
      attendanceData.push(["SECTION BREAKDOWN & VERIFICATION SUMMARY"]);
      sectionPositions.forEach((pos) => {
        const count = sectionGroups[pos]?.length || 0;
        const branchSample = sectionGroups[pos]?.[0]?.student.branch || "General";
        attendanceData.push([`Section ${pos} Total Candidates:`, count, `Primary Branch: ${branchSample}`]);
      });
      attendanceData.push([]);
      attendanceData.push(["Chief Superintendent:", collegeProfile?.chiefSuperintendent || "Prof. S. K. Narayan"]);
      attendanceData.push(["Invigilator In-Charge:", "__________________________ (Signature)"]);

      const wsAttendance = XLSX.utils.aoa_to_sheet(attendanceData);
      XLSX.utils.book_append_sheet(wb, wsAttendance, `Att - Room ${room.roomNumber}`);
    });

    const filePrefix = activeSession?.examMode === "MST" ? "MST_SeatingChart" : "EndSem_SeatingChart";
    XLSX.writeFile(wb, `${filePrefix}_${examDate.replace(/-/g, "_")}.xlsx`);
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
