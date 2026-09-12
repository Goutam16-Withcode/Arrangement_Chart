"use client";
import React, { createContext, useContext, useState, useEffect } from "react";
import {
  RoomConfig,
  Student,
  Invigilator,
  RoomSeating,
  SeatingMetrics,
  CollegeProfile,
  ExamSession,
  ExamMode,
} from "../types";
import {
  SAMPLE_ROOMS,
  SAMPLE_INVIGILATORS,
  DEFAULT_COLLEGE_PROFILE,
  EXAM_SESSIONS,
  generateStudentsForSession,
} from "../sampleData";
import { SeatingEngine } from "../seatingEngine";

interface SeatingContextType {
  collegeProfile: CollegeProfile;
  updateCollegeProfile: (profile: Partial<CollegeProfile>) => void;
  examMode: ExamMode;
  setExamMode: (mode: ExamMode) => void;
  activeSession: ExamSession;
  sessions: ExamSession[];
  switchSession: (sessionId: string) => void;
  updateActiveSession: (update: Partial<ExamSession>) => void;
  addCustomSession: (newSession: ExamSession) => void;
  isConfigModalOpen: boolean;
  setIsConfigModalOpen: (open: boolean) => void;
  rooms: RoomConfig[];
  students: Student[];
  invigilators: Invigilator[];
  roomSeatings: RoomSeating[];
  metrics: SeatingMetrics;
  selectedRoomNumber: string;
  setSelectedRoomNumber: (roomNum: string) => void;
  recalculateSeating: (customRooms?: RoomConfig[], customStudents?: Student[]) => void;
  swapSeats: (roomNumber: string, seatIdA: string, seatIdB: string) => void;
  toggleAttendance: (roomNumber: string, serialNo: number) => void;
  addCustomRoom: (room: RoomConfig) => void;
  setStudentsList: (students: Student[]) => void;
  resetToDefaults: () => void;
}

const SeatingContext = createContext<SeatingContextType | undefined>(undefined);

export const SeatingProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [collegeProfile, setCollegeProfile] = useState<CollegeProfile>(DEFAULT_COLLEGE_PROFILE);
  const [sessions, setSessions] = useState<ExamSession[]>(EXAM_SESSIONS);
  const [examMode, setExamModeState] = useState<ExamMode>("END_SEM");
  const [activeSession, setActiveSession] = useState<ExamSession>(EXAM_SESSIONS[2]); // Default End-Sem Slot 1
  const [isConfigModalOpen, setIsConfigModalOpen] = useState(false);

  const [rooms, setRooms] = useState<RoomConfig[]>(SAMPLE_ROOMS);
  const [students, setStudents] = useState<Student[]>([]);
  const [invigilators] = useState<Invigilator[]>(SAMPLE_INVIGILATORS);
  const [selectedRoomNumber, setSelectedRoomNumber] = useState<string>("302");

  const [roomSeatings, setRoomSeatings] = useState<RoomSeating[]>([]);
  const [metrics, setMetrics] = useState<SeatingMetrics>({
    totalRooms: 0,
    totalCapacity: 0,
    totalStudents: 0,
    seatedStudents: 0,
    utilizationRate: 0,
    conflictFreeRate: 100,
    branchesCount: {},
  });

  // Hydrate from localStorage if available
  useEffect(() => {
    try {
      const savedProfile = localStorage.getItem("smart_seating_college_profile");
      if (savedProfile) {
        setCollegeProfile(JSON.parse(savedProfile));
      }
    } catch {
      // ignore
    }
  }, []);

  // Load session students & calculate seating on mount
  useEffect(() => {
    const sessionStudents = generateStudentsForSession(activeSession);
    setStudents(sessionStudents);
    const result = SeatingEngine.generateSeating(SAMPLE_ROOMS, sessionStudents);
    setRoomSeatings(result.roomSeatings);
    setMetrics(result.metrics);
  }, []);

  const switchSession = (sessionId: string) => {
    const found = sessions.find((s) => s.id === sessionId);
    if (!found) return;

    setActiveSession(found);
    setExamModeState(found.examMode);

    const sessionStudents = generateStudentsForSession(found);
    setStudents(sessionStudents);

    const result = SeatingEngine.generateSeating(rooms, sessionStudents);
    setRoomSeatings(result.roomSeatings);
    setMetrics(result.metrics);
  };

  const setExamMode = (mode: ExamMode) => {
    setExamModeState(mode);
    const matchingSession = sessions.find((s) => s.examMode === mode) || sessions[0];
    switchSession(matchingSession.id);
  };

  const updateCollegeProfile = (profileUpdate: Partial<CollegeProfile>) => {
    setCollegeProfile((prev) => {
      const updated = { ...prev, ...profileUpdate };
      try {
        localStorage.setItem("smart_seating_college_profile", JSON.stringify(updated));
      } catch {
        // ignore
      }
      return updated;
    });
  };

  const updateActiveSession = (update: Partial<ExamSession>) => {
    const updated = { ...activeSession, ...update };
    setActiveSession(updated);
    setSessions((prev) =>
      prev.map((s) => (s.id === activeSession.id ? { ...s, ...update } : s))
    );
  };

  const addCustomSession = (newSession: ExamSession) => {
    setSessions((prev) => [...prev, newSession]);
    switchSession(newSession.id);
  };

  const recalculateSeating = (customRooms?: RoomConfig[], customStudents?: Student[]) => {
    const activeRooms = customRooms || rooms;
    const activeStudents = customStudents || students;
    const result = SeatingEngine.generateSeating(activeRooms, activeStudents);
    setRoomSeatings(result.roomSeatings);
    setMetrics(result.metrics);
  };

  const swapSeats = (roomNumber: string, seatIdA: string, seatIdB: string) => {
    setRoomSeatings((prev) =>
      prev.map((rs) => {
        if (rs.roomConfig.roomNumber === roomNumber) {
          const updatedSeats = SeatingEngine.swapSeats(rs.seats, rs.roomConfig, seatIdA, seatIdB);
          return {
            ...rs,
            seats: updatedSeats,
          };
        }
        return rs;
      })
    );
  };

  const toggleAttendance = (roomNumber: string, serialNo: number) => {
    setRoomSeatings((prev) =>
      prev.map((rs) => {
        if (rs.roomConfig.roomNumber === roomNumber) {
          const updatedList = rs.attendanceList.map((item) =>
            item.serialNo === serialNo ? { ...item, present: !item.present } : item
          );
          return {
            ...rs,
            attendanceList: updatedList,
          };
        }
        return rs;
      })
    );
  };

  const addCustomRoom = (room: RoomConfig) => {
    const updated = [...rooms, room];
    setRooms(updated);
    recalculateSeating(updated, students);
  };

  const setStudentsList = (newStudents: Student[]) => {
    setStudents(newStudents);
    recalculateSeating(rooms, newStudents);
  };

  const resetToDefaults = () => {
    setCollegeProfile(DEFAULT_COLLEGE_PROFILE);
    setSessions(EXAM_SESSIONS);
    setRooms(SAMPLE_ROOMS);
    try {
      localStorage.removeItem("smart_seating_college_profile");
    } catch {
      // ignore
    }
    switchSession("endsem-slot-1");
    setSelectedRoomNumber("302");
  };

  return (
    <SeatingContext.Provider
      value={{
        collegeProfile,
        updateCollegeProfile,
        examMode,
        setExamMode,
        activeSession,
        sessions,
        switchSession,
        updateActiveSession,
        addCustomSession,
        isConfigModalOpen,
        setIsConfigModalOpen,
        rooms,
        students,
        invigilators,
        roomSeatings,
        metrics,
        selectedRoomNumber,
        setSelectedRoomNumber,
        recalculateSeating,
        swapSeats,
        toggleAttendance,
        addCustomRoom,
        setStudentsList,
        resetToDefaults,
      }}
    >
      {children}
    </SeatingContext.Provider>
  );
};

export const useSeating = () => {
  const context = useContext(SeatingContext);
  if (!context) {
    throw new Error("useSeating must be used within a SeatingProvider");
  }
  return context;
};

