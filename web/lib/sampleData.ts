import { RoomConfig, Student, Invigilator } from "./types";

export const SAMPLE_ROOMS: RoomConfig[] = [
  {
    roomNumber: "302",
    building: "Academic Block A",
    floor: 3,
    rows: 5,
    benchesPerRow: 4,
    studentsPerBench: 3,
    leftBranchName: "B.Tech CSE - Y2",
    middleBranchName: "B.Tech ME - Y3",
    rightBranchName: "B.Tech ECE - Y4",
    customName: "Computing Lecture Hall 302",
  },
  {
    roomNumber: "304",
    building: "Academic Block A",
    floor: 3,
    rows: 4,
    benchesPerRow: 3,
    studentsPerBench: 3,
    leftBranchName: "B.Tech CSE - Y2",
    middleBranchName: "B.Tech ME - Y3",
    rightBranchName: "B.Tech ECE - Y4",
    customName: "Seminar Room 304",
  },
  {
    roomNumber: "101",
    building: "Engineering Block B",
    floor: 1,
    rows: 6,
    benchesPerRow: 4,
    studentsPerBench: 2,
    leftBranchName: "B.Tech IT - Y2",
    middleBranchName: "B.Tech Civil - Y3",
    customName: "Drawing Hall 101",
  },
];

export const generateSampleStudents = (): Student[] => {
  const branches = [
    { code: "CSE", name: "Computer Science", subject: "CS401 - Database Systems", year: 2 },
    { code: "ME", name: "Mechanical Engineering", subject: "ME502 - Thermodynamics", year: 3 },
    { code: "ECE", name: "Electronics & Comm.", subject: "EC603 - Signal Processing", year: 4 },
    { code: "IT", name: "Information Tech", subject: "IT404 - Web Technologies", year: 2 },
  ];

  const firstNames = ["Aarav", "Aditi", "Ananya", "Dev", "Diya", "Ishaan", "Kavya", "Manish", "Neha", "Pranav", "Pooja", "Rahul", "Rhea", "Rohan", "Sanya", "Shaurya", "Sneha", "Tanvi", "Varun", "Yash", "Aditya", "Bhavna", "Chetan", "Deepika", "Gautam", "Harsh", "Isha", "Jatin", "Kiran", "Meera", "Nikhil", "Om", "Priya", "Raj", "Simran", "Tushar", "Urvi", "Vikas", "Zoya"];
  const lastNames = ["Sharma", "Verma", "Patel", "Reddy", "Gupta", "Mehta", "Singh", "Nair", "Iyer", "Chopra", "Joshi", "Bhat", "Rao", "Mishra", "Malhotra", "Saxena", "Kapoor", "Pandey"];

  const students: Student[] = [];
  let idCounter = 1;

  branches.forEach((b) => {
    const count = 45; // 45 students per branch
    for (let i = 1; i <= count; i++) {
      const fn = firstNames[(idCounter * 7) % firstNames.length];
      const ln = lastNames[(idCounter * 11) % lastNames.length];
      const rollNum = `${b.code}2026${String(i).padStart(3, "0")}`;

      students.push({
        id: `std-${idCounter}`,
        rollNo: rollNum,
        name: `${fn} ${ln}`,
        branch: b.code,
        year: b.year,
        subjectCode: b.subject.split(" - ")[0],
        subjectName: b.subject.split(" - ")[1],
        photoUrl: `https://images.unsplash.com/photo-${1534528741775 + (idCounter % 50)}?w=100&auto=format&fit=crop&q=80`,
        hasSpecialNeed: idCounter % 35 === 0,
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
