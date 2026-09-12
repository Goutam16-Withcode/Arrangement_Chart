import { RoomConfig, Student, Invigilator, CollegeProfile, ExamSession } from "./types";

export const DEFAULT_COLLEGE_PROFILE: CollegeProfile = {
  collegeName: "Apex University of Technology & Management",
  collegeCode: "AUTM-2026",
  academicYear: "Academic Session 2025-2026",
  semester: "Semester Examination / Term Assessment",
  examCenterCode: "EXAM-CTR-01",
  chiefSuperintendent: "Dr. Controller of Examinations",
};

export const EXAM_SESSIONS: ExamSession[] = [
  {
    id: "mst-slot-1",
    examMode: "MST",
    title: "Mid-Semester Assessment (MST) • Slot 1",
    date: "Day 1 (Morning Session)",
    timing: "10:00 AM - 11:30 AM (1.5 Hours)",
    slot: "Slot 1 (Morning)",
    courses: [
      { branch: "CSE", year: 2, code: "CS401", name: "Database Management Systems" },
      { branch: "ME", year: 3, code: "ME502", name: "Thermodynamics & Heat Transfer" },
      { branch: "ECE", year: 4, code: "EC603", name: "Digital Signal Processing" },
    ],
  },
  {
    id: "mst-slot-2",
    examMode: "MST",
    title: "Mid-Semester Assessment (MST) • Slot 2",
    date: "Day 1 (Afternoon Session)",
    timing: "02:30 PM - 04:00 PM (1.5 Hours)",
    slot: "Slot 2 (Afternoon)",
    courses: [
      { branch: "CSE", year: 3, code: "CS601", name: "Operating Systems & Kernels" },
      { branch: "ME", year: 2, code: "ME401", name: "Fluid Mechanics & Machinery" },
      { branch: "ECE", year: 3, code: "EC501", name: "Microprocessors & Embedded Sys" },
    ],
  },
  {
    id: "endsem-slot-1",
    examMode: "END_SEM",
    title: "End-Semester Major Theory Examination • Slot 1",
    date: "Day 1 (Morning Session)",
    timing: "09:30 AM - 12:30 PM (3.0 Hours)",
    slot: "Slot 1 (Morning)",
    courses: [
      { branch: "CSE", year: 2, code: "CS401", name: "Design & Analysis of Algorithms" },
      { branch: "ME", year: 3, code: "ME602", name: "Kinematics & Dynamics of Machines" },
      { branch: "ECE", year: 4, code: "EC803", name: "Wireless Networks & IoT" },
    ],
  },
  {
    id: "endsem-slot-2",
    examMode: "END_SEM",
    title: "End-Semester Major Theory Examination • Slot 2",
    date: "Day 1 (Afternoon Session)",
    timing: "02:00 PM - 05:00 PM (3.0 Hours)",
    slot: "Slot 2 (Afternoon)",
    courses: [
      { branch: "CSE", year: 4, code: "CS801", name: "Distributed Systems & Cloud" },
      { branch: "ME", year: 4, code: "ME802", name: "Industrial Automation & Robotics" },
      { branch: "ECE", year: 2, code: "EC402", name: "Analog Electronic Circuits" },
    ],
  },
];

export const SAMPLE_ROOMS: RoomConfig[] = [
  {
    roomNumber: "302",
    building: "Academic Block A",
    floor: 3,
    rows: 5,
    benchesPerRow: 4,
    studentsPerBench: 3,
    leftBranchName: "Branch 1",
    middleBranchName: "Branch 2",
    rightBranchName: "Branch 3",
    customName: "Lecture Hall 302",
  },
  {
    roomNumber: "304",
    building: "Academic Block A",
    floor: 3,
    rows: 4,
    benchesPerRow: 3,
    studentsPerBench: 3,
    leftBranchName: "Branch 1",
    middleBranchName: "Branch 2",
    rightBranchName: "Branch 3",
    customName: "Seminar Room 304",
  },
  {
    roomNumber: "101",
    building: "Main Complex B",
    floor: 1,
    rows: 6,
    benchesPerRow: 4,
    studentsPerBench: 2,
    leftBranchName: "Branch 1",
    middleBranchName: "Branch 2",
    customName: "Exam Hall 101",
  },
];

export const generateStudentsForSession = (session: ExamSession): Student[] => {
  const firstNames = [
    "Aarav", "Aditi", "Ananya", "Dev", "Diya", "Ishaan", "Kavya", "Manish",
    "Neha", "Pranav", "Pooja", "Rahul", "Rhea", "Rohan", "Sanya", "Shaurya",
    "Sneha", "Tanvi", "Varun", "Yash", "Aditya", "Bhavna", "Chetan", "Deepika",
    "Gautam", "Harsh", "Isha", "Jatin", "Kiran", "Meera", "Nikhil", "Om",
    "Priya", "Raj", "Simran", "Tushar", "Urvi", "Vikas", "Zoya", "Aman",
    "Kunal", "Mehak", "Ravi", "Payal", "Sameer", "Tarun", "Divya", "Siddharth"
  ];
  const lastNames = [
    "Sharma", "Verma", "Patel", "Reddy", "Gupta", "Mehta", "Singh", "Nair",
    "Iyer", "Chopra", "Joshi", "Bhat", "Rao", "Mishra", "Malhotra", "Saxena"
  ];

  const students: Student[] = [];
  let idCounter = 1;

  session.courses.forEach((course) => {
    const count = 40;
    for (let i = 1; i <= count; i++) {
      const fn = firstNames[(idCounter * 7) % firstNames.length];
      const ln = lastNames[(idCounter * 11) % lastNames.length];
      const rollNum = `${course.branch}2026${String(i).padStart(3, "0")}`;

      students.push({
        id: `std-${session.id}-${idCounter}`,
        rollNo: rollNum,
        name: `${fn} ${ln}`,
        branch: course.branch,
        year: course.year,
        subjectCode: course.code,
        subjectName: course.name,
        photoUrl: `https://images.unsplash.com/photo-${1534528741775 + (idCounter % 50)}?w=100&auto=format&fit=crop&q=80`,
        hasSpecialNeed: idCounter % 32 === 0,
      });
      idCounter++;
    }
  });

  return students;
};

export const SAMPLE_INVIGILATORS: Invigilator[] = [
  {
    id: 1,
    name: "Dr. Rajesh Kulkarni",
    department: "Computer Science",
    designation: "Associate Professor",
    image: "https://images.unsplash.com/photo-1534528741775-53994a69daeb?w=100&auto=format&fit=crop&q=80",
    assignedRoom: "302",
    status: "Active",
  },
  {
    id: 2,
    name: "Prof. Sunita Rao",
    department: "Mechanical Engg",
    designation: "Assistant Professor",
    image: "https://images.unsplash.com/photo-1580489944761-15a19d654956?w=100&auto=format&fit=crop&q=80",
    assignedRoom: "304",
    status: "Active",
  },
  {
    id: 3,
    name: "Dr. Vikram Sengupta",
    department: "Applied Mathematics",
    designation: "Senior Faculty",
    image: "https://images.unsplash.com/photo-1507003211169-0a1dd7228f2d?w=100&auto=format&fit=crop&q=80",
    assignedRoom: "101",
    status: "Active",
  },
  {
    id: 4,
    name: "Dr. Meenakshi Sundaram",
    department: "Electronics & Comm",
    designation: "Professor & HOD",
    image: "https://images.unsplash.com/photo-1544005313-94ddf0286df2?w=100&auto=format&fit=crop&q=80",
    status: "Standby",
  },
];
