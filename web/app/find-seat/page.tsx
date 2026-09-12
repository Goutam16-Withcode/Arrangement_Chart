"use client";
import React, { useState } from "react";
import { useSeating } from "@/lib/context/SeatingContext";
import { Student, RoomConfig } from "@/lib/types";
import { PinContainer } from "@/components/ui/3d-pin";
import { BackgroundGradient } from "@/components/ui/background-gradient";
import { QRCodeSVG } from "qrcode.react";
import {
  Search,
  Sparkles,
  MapPin,
  CheckCircle2,
  AlertCircle,
} from "lucide-react";

interface FoundAllocation {
  student: Student;
  room: RoomConfig;
  benchIndex: number;
  positionLabel: string;
  rowIndex: number;
}

export default function FindSeatPage() {
  const { roomSeatings } = useSeating();
  const [query, setQuery] = useState("CSE2026001");
  const [searchedRoll, setSearchedRoll] = useState("CSE2026001");

  // Search through all rooms to find matching seat
  let foundAllocation: FoundAllocation | null = null;

  for (const rs of roomSeatings) {
    for (const bench of rs.seats) {
      for (const seat of bench) {
        if (
          seat.student &&
          (seat.student.rollNo.toLowerCase() === searchedRoll.toLowerCase() ||
            seat.student.name.toLowerCase().includes(searchedRoll.toLowerCase()))
        ) {
          foundAllocation = {
            student: seat.student,
            room: rs.roomConfig,
            benchIndex: seat.benchIndex + 1,
            positionLabel: seat.positionLabel,
            rowIndex: seat.rowIndex + 1,
          };
          break;
        }
      }
      if (foundAllocation) break;
    }
    if (foundAllocation) break;
  }

  const handleSearch = (e: React.FormEvent) => {
    e.preventDefault();
    if (query.trim()) {
      setSearchedRoll(query.trim());
    }
  };

  const allocation: FoundAllocation | null = foundAllocation;

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-10 space-y-10 bg-[#FBF9F4]">
      {/* Header */}
      <div className="text-center max-w-2xl mx-auto space-y-3">
        <div className="inline-flex items-center gap-2 px-3.5 py-1 rounded-full bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-semibold">
          <Sparkles className="h-3.5 w-3.5 text-emerald-600" />
          <span>Student Self-Service Portal</span>
        </div>
        <h1 className="text-3xl sm:text-4xl font-extrabold text-slate-900 tracking-tight">
          Find Your Exam Seat & Digital Pass
        </h1>
        <p className="text-xs text-slate-500">
          Enter your university Roll Number to instantly locate your examination hall, bench, and digital entry pass.
        </p>

        {/* Search Bar */}
        <form onSubmit={handleSearch} className="pt-4 flex items-center gap-2 max-w-md mx-auto">
          <div className="relative flex-1">
            <Search className="h-4 w-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              placeholder="e.g. CSE2026001, ME2026012, or Name..."
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              className="w-full pl-10 pr-4 py-3 rounded-2xl bg-white border border-[#E0D9CB] text-xs font-mono text-slate-900 placeholder-slate-400 focus:outline-none focus:border-emerald-500 transition shadow-xs"
            />
          </div>
          <button
            type="submit"
            className="px-5 py-3 rounded-2xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs tracking-wider uppercase transition shadow-md shadow-emerald-600/20"
          >
            Locate
          </button>
        </form>
      </div>

      {/* Result Section */}
      {allocation !== null ? (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-10 items-center max-w-4xl mx-auto pt-6">
          {/* Card 1: 3D Campus Pin Map */}
          <div className="flex items-center justify-center min-h-[22rem]">
            <PinContainer
              title={`${allocation.room.building} • Room ${allocation.room.roomNumber}`}
              href="/studio"
            >
              <div className="flex basis-full flex-col p-4 tracking-tight text-slate-900 sm:basis-1/2 w-[18rem] h-[18rem]">
                <h3 className="max-w-xs !pb-2 !m-0 font-bold text-base text-slate-900">
                  Room {allocation.room.roomNumber}
                </h3>
                <div className="text-xs !m-0 !p-0 font-normal text-slate-500">
                  Floor {allocation.room.floor} • {allocation.room.building}
                </div>
                <div className="flex flex-1 w-full rounded-xl mt-4 bg-emerald-50 border border-emerald-200 flex-col items-center justify-center p-4 text-center">
                  <MapPin className="h-8 w-8 text-emerald-600 animate-bounce mb-2" />
                  <div className="text-xs font-mono font-bold text-slate-900">
                    Row #{allocation.rowIndex} • Bench #{allocation.benchIndex}
                  </div>
                  <div className="text-[10px] text-emerald-800 font-mono mt-1 px-2.5 py-0.5 rounded-full bg-white border border-emerald-300 font-bold shadow-2xs">
                    Seat Tag: {allocation.positionLabel} (
                    {allocation.positionLabel === "F-1"
                      ? "Left"
                      : allocation.positionLabel === "S-1"
                      ? "Middle"
                      : "Right"}
                    )
                  </div>
                </div>
              </div>
            </PinContainer>
          </div>

          {/* Card 2: Digital Admit Ticket Pass (BackgroundGradient) */}
          <div>
            <BackgroundGradient className="rounded-[22px] p-6 bg-white border border-[#E8E2D4] shadow-sm space-y-5">
              <div className="flex items-center justify-between border-b border-[#E8E2D4] pb-4">
                <div>
                  <div className="text-[10px] font-mono text-emerald-700 uppercase tracking-widest font-bold">
                    Digital Exam Pass
                  </div>
                  <h2 className="text-xl font-bold text-slate-900 tracking-tight mt-0.5">
                    {allocation.student.name}
                  </h2>
                  <div className="text-xs font-mono text-slate-500">
                    {allocation.student.rollNo}
                  </div>
                </div>

                {/* QR Code */}
                <div className="p-2 bg-white rounded-xl shadow-xs border border-[#E8E2D4]">
                  <QRCodeSVG
                    value={`EXAM-VERIFY:${allocation.student.rollNo}:ROOM-${allocation.room.roomNumber}:BENCH-${allocation.benchIndex}`}
                    size={64}
                  />
                </div>
              </div>

              {/* Exam details grid */}
              <div className="grid grid-cols-2 gap-3 text-xs">
                <div className="p-2.5 rounded-xl bg-[#FAF8F3] border border-[#E8E2D4]">
                  <span className="text-slate-500 block text-[10px] uppercase font-mono">
                    Department
                  </span>
                  <span className="font-bold text-slate-800">
                    {allocation.student.branch} (Year {allocation.student.year})
                  </span>
                </div>

                <div className="p-2.5 rounded-xl bg-[#FAF8F3] border border-[#E8E2D4]">
                  <span className="text-slate-500 block text-[10px] uppercase font-mono">
                    Subject Code
                  </span>
                  <span className="font-bold text-slate-800 font-mono">
                    {allocation.student.subjectCode}
                  </span>
                </div>

                <div className="p-2.5 rounded-xl bg-[#FAF8F3] border border-[#E8E2D4]">
                  <span className="text-slate-500 block text-[10px] uppercase font-mono">
                    Assigned Hall
                  </span>
                  <span className="font-bold text-emerald-700 font-mono">
                    Room {allocation.room.roomNumber}
                  </span>
                </div>

                <div className="p-2.5 rounded-xl bg-[#FAF8F3] border border-[#E8E2D4]">
                  <span className="text-slate-500 block text-[10px] uppercase font-mono">
                    Seat Coordinate
                  </span>
                  <span className="font-bold text-teal-800 font-mono">
                    Bench {allocation.benchIndex} ({allocation.positionLabel})
                  </span>
                </div>
              </div>

              {/* Verification Stamp */}
              <div className="p-3 rounded-xl bg-emerald-50 border border-emerald-200 flex items-center justify-between text-xs text-emerald-800 font-semibold">
                <div className="flex items-center gap-2">
                  <CheckCircle2 className="h-4 w-4 text-emerald-600" />
                  <span>Verified Entry Ticket</span>
                </div>
                <span className="font-mono text-[10px]">Session: 09:30 AM</span>
              </div>
            </BackgroundGradient>
          </div>
        </div>
      ) : (
        <div className="max-w-md mx-auto p-8 rounded-3xl bg-white border border-[#E8E2D4] text-center space-y-3 shadow-sm">
          <AlertCircle className="h-8 w-8 text-amber-500 mx-auto" />
          <h3 className="text-base font-bold text-slate-900">No Matching Record Found</h3>
          <p className="text-xs text-slate-500">
            Could not find student matching &quot;{searchedRoll}&quot;. Try searching with demo roll number{" "}
            <code className="text-emerald-700 font-bold">CSE2026001</code> or <code className="text-emerald-700 font-bold">ME2026005</code>.
          </p>
        </div>
      )}
    </div>
  );
}
