export type ExamMode = "MST" | "END_SEM";

export interface CollegeProfile {
  collegeName: string;
  collegeCode: string;
  academicYear: string;
  semester: string;
  examCenterCode: string;
  chiefSuperintendent: string;
}

export interface ExamSession {
  id: string;
  examMode: ExamMode;
  title: string;
  date: string;
  timing: string;
  slot: string;
  courses: {
    branch: string;
    year: number;
    code: string;
    name: string;
  }[];
}

export interface Student {
  id: string;
  rollNo: string;
  name: string;
  branch: string;
  year: number;
  subjectCode: string;
  subjectName: string;
  photoUrl?: string;
  hasSpecialNeed?: boolean;
}

export interface SeatSlot {
  id: string;
  roomNumber: string;
  benchIndex: number;
  positionIndex: number; // 0 for F-1 (Left), 1 for S-1 (Middle), 2 for T-1 (Right), etc.
  positionLabel: string; // 'F-1', 'S-1', 'T-1', etc.
  rowIndex: number;
  colIndex: number;
  student: Student | null;
  isDefective?: boolean;
  hasConflict?: boolean;
  conflictReason?: string;
}

export interface RoomConfig {
  roomNumber: string;
  building: string;
  floor: number;
  rows: number;
  benchesPerRow: number;
  studentsPerBench: number; // 1, 2, 3, 4
  leftBranchName?: string;
  middleBranchName?: string;
  rightBranchName?: string;
  customName?: string;
}

export interface RoomSeating {
  roomConfig: RoomConfig;
  totalCapacity: number;
  assignedCount: number;
  seats: SeatSlot[][]; // Matrix of [benchRow][seatPosition]
  attendanceList: {
    serialNo: number;
    position: string;
    student: Student;
    present: boolean;
  }[];
  invigilator?: Invigilator;
}

export interface Invigilator {
  id: number;
  name: string;
  department: string;
  designation: string;
  image: string;
  assignedRoom?: string;
  status: "Assigned" | "Standby" | "Active";
}

export interface SeatingMetrics {
  totalRooms: number;
  totalCapacity: number;
  totalStudents: number;
  seatedStudents: number;
  utilizationRate: number;
  conflictFreeRate: number;
  branchesCount: { [branch: string]: number };
}
