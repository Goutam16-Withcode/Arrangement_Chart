"use client";
import React, { createContext, useContext, useState, useEffect } from "react";
import { RoomConfig, Student, Invigilator, RoomSeating, SeatingMetrics } from "../types";
import { SAMPLE_ROOMS, generateSampleStudents, SAMPLE_INVIGILATORS } from "../sampleData";
import { SeatingEngine } from "../seatingEngine";

interface SeatingContextType {
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

  // Initialize with sample data on mount
  useEffect(() => {
    const sampleStudents = generateSampleStudents();
    setStudents(sampleStudents);
    const result = SeatingEngine.generateSeating(SAMPLE_ROOMS, sampleStudents);
    setRoomSeatings(result.roomSeatings);
    setMetrics(result.metrics);
  }, []);

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
    const sampleStudents = generateSampleStudents();
    setRooms(SAMPLE_ROOMS);
    setStudents(sampleStudents);
    const result = SeatingEngine.generateSeating(SAMPLE_ROOMS, sampleStudents);
    setRoomSeatings(result.roomSeatings);
    setMetrics(result.metrics);
    setSelectedRoomNumber("302");
  };

  return (
    <SeatingContext.Provider
      value={{
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
