"use client";
import React, { useState } from "react";
import { useSeating } from "@/lib/context/SeatingContext";
import { RoomConfig } from "@/lib/types";
import {
  Plus,
  Sliders,
  Sparkles,
} from "lucide-react";
import { useRouter } from "next/navigation";
import confetti from "canvas-confetti";

export default function BuilderPage() {
  const { addCustomRoom, setSelectedRoomNumber, collegeProfile } = useSeating();
  const router = useRouter();

  const [roomNumber, setRoomNumber] = useState("401");
  const [building, setBuilding] = useState("Technology Tower");
  const [floor, setFloor] = useState(4);
  const [rows, setRows] = useState(5);
  const [benchesPerRow, setBenchesPerRow] = useState(4);
  const [studentsPerBench, setStudentsPerBench] = useState(3);
  const [customName, setCustomName] = useState("Innovation Lab 401");

  const totalBenches = rows * benchesPerRow;
  const totalCapacity = totalBenches * studentsPerBench;

  const handleSaveRoom = (e: React.FormEvent) => {
    e.preventDefault();

    const newRoom: RoomConfig = {
      roomNumber,
      building,
      floor,
      rows,
      benchesPerRow,
      studentsPerBench,
      customName,
      leftBranchName: "Branch 1",
      middleBranchName: "Branch 2",
      rightBranchName: "Branch 3",
    };

    addCustomRoom(newRoom);
    setSelectedRoomNumber(roomNumber);

    confetti({
      particleCount: 40,
      spread: 70,
      origin: { y: 0.6 },
      colors: ["#D4F754", "#B8B5FF", "#161618"],
    });

    router.push("/studio");
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div>
        <div className="flex items-center gap-2 text-xs font-mono font-bold uppercase tracking-wider mb-1 flex-wrap text-slate-800">
          <span>✦ {collegeProfile.collegeName}</span>
          <span className="text-slate-400">•</span>
          <span>Room Geometry Designer</span>
        </div>
        <h1 className="text-2xl sm:text-3xl font-black text-slate-900 tracking-tight">
          2D Visual Room Blueprint Builder
        </h1>
        <p className="text-xs text-slate-500 mt-1 font-medium">
          Configure custom classroom, auditorium, or computer lab layouts with dynamic bench sizes.
        </p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* Controls Configuration Form */}
        <div className="lg:col-span-5 bento-card p-6 space-y-6">
          <form onSubmit={handleSaveRoom} className="space-y-4">
            <h3 className="text-sm font-black text-slate-900 uppercase tracking-wider font-mono flex items-center gap-2">
              <Sliders className="h-4 w-4 text-slate-900" />
              <span>Room Specifications</span>
            </h3>

            {/* Room Number & Floor */}
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">
                  Room Number *
                </label>
                <input
                  type="text"
                  required
                  value={roomNumber}
                  onChange={(e) => setRoomNumber(e.target.value)}
                  className="w-full px-3.5 py-2.5 rounded-xl bg-slate-100 border border-slate-200 text-xs text-slate-900 focus:outline-none focus:border-black font-bold"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">
                  Floor Level
                </label>
                <input
                  type="number"
                  min={0}
                  value={floor}
                  onChange={(e) => setFloor(Number(e.target.value))}
                  className="w-full px-3.5 py-2.5 rounded-xl bg-slate-100 border border-slate-200 text-xs text-slate-900 focus:outline-none focus:border-black font-bold"
                />
              </div>
            </div>

            <div>
              <label className="block text-xs font-bold text-slate-700 mb-1">
                Building / Block Name
              </label>
              <input
                type="text"
                value={building}
                onChange={(e) => setBuilding(e.target.value)}
                className="w-full px-3.5 py-2.5 rounded-xl bg-slate-100 border border-slate-200 text-xs text-slate-900 focus:outline-none focus:border-black font-bold"
              />
            </div>

            {/* Dimension Sliders */}
            <div className="space-y-3 pt-2">
              <div>
                <div className="flex justify-between text-xs font-bold text-slate-700 mb-1">
                  <span>Number of Rows</span>
                  <span className="text-black font-mono font-black bg-[#D4F754] px-2 py-0.5 rounded-md">{rows} Rows</span>
                </div>
                <input
                  type="range"
                  min="2"
                  max="12"
                  value={rows}
                  onChange={(e) => setRows(Number(e.target.value))}
                  className="w-full accent-black"
                />
              </div>

              <div>
                <div className="flex justify-between text-xs font-bold text-slate-700 mb-1">
                  <span>Benches per Row</span>
                  <span className="text-black font-mono font-black bg-[#B8B5FF] px-2 py-0.5 rounded-md">{benchesPerRow} Benches</span>
                </div>
                <input
                  type="range"
                  min="1"
                  max="8"
                  value={benchesPerRow}
                  onChange={(e) => setBenchesPerRow(Number(e.target.value))}
                  className="w-full accent-black"
                />
              </div>

              <div>
                <div className="flex justify-between text-xs font-bold text-slate-700 mb-1">
                  <span>Students per Bench</span>
                  <span className="text-slate-900 font-mono font-bold">
                    {studentsPerBench} Seater (
                    {studentsPerBench === 1
                      ? "F-1"
                      : studentsPerBench === 2
                      ? "F-1, S-1"
                      : studentsPerBench === 3
                      ? "F-1, S-1, T-1"
                      : "F-1, S-1, T-1, F-2"}
                    )
                  </span>
                </div>
                <div className="grid grid-cols-4 gap-2">
                  {[1, 2, 3, 4].map((n) => (
                    <button
                      key={n}
                      type="button"
                      onClick={() => setStudentsPerBench(n)}
                      className={`py-2 rounded-xl text-xs font-mono font-bold transition border ${
                        studentsPerBench === n
                          ? "bg-[#161618] text-[#D4F754] border-[#161618] shadow-sm"
                          : "bg-slate-100 text-slate-700 border-slate-200 hover:text-slate-900"
                      }`}
                    >
                      {n} {n === 1 ? "Seat" : "Seats"}
                    </button>
                  ))}
                </div>
              </div>
            </div>

            {/* Capacity Banner */}
            <div className="p-4 rounded-2xl bg-slate-100 border border-slate-200 flex items-center justify-between text-xs">
              <div>
                <div className="text-slate-500 font-bold">Calculated Capacity</div>
                <div className="text-lg font-black font-mono text-slate-900">
                  {totalCapacity} Student Desks
                </div>
              </div>
              <div className="text-right text-[11px] text-black bg-[#D4F754] px-2.5 py-1 rounded-full font-mono font-black">
                {totalBenches} Total Benches
              </div>
            </div>

            <button
              type="submit"
              className="w-full py-3.5 rounded-full bg-[#161618] hover:bg-black text-white font-black text-xs tracking-wider uppercase transition shadow-md flex items-center justify-center gap-2 ring-2 ring-[#D4F754]/30"
            >
              <Plus className="h-4 w-4 text-[#D4F754]" />
              <span>Add Room & Generate Seating</span>
            </button>
          </form>
        </div>

        {/* Real-time 2D Blueprint Live Preview */}
        <div className="lg:col-span-7 bg-white border border-slate-200/80 p-6 rounded-[28px] shadow-xs space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="text-sm font-black text-slate-900 uppercase tracking-wider font-mono">
              Live Blueprint Preview
            </h3>
            <span className="text-xs px-3 py-1 rounded-full bg-[#161618] text-[#D4F754] font-mono font-bold">
              Room {roomNumber} • {totalCapacity} Seats
            </span>
          </div>

          {/* Blackboard Stage */}
          <div className="py-2.5 px-4 rounded-xl bg-[#161618] text-center text-xs font-black text-white font-mono shadow-sm">
            🖥️ Front Screen / Invigilator Podium
          </div>

          {/* Rendered Live Blueprint Grid */}
          <div className="space-y-3 max-h-[440px] overflow-y-auto no-visible-scrollbar p-4 rounded-2xl bg-slate-50 border border-slate-200">
            {Array.from({ length: rows }).map((_, r) => (
              <div key={r} className="space-y-1">
                <div className="text-[10px] text-slate-500 font-mono font-bold">Row {r + 1}</div>
                <div
                  className="grid gap-2"
                  style={{
                    gridTemplateColumns: `repeat(${benchesPerRow}, minmax(0, 1fr))`,
                  }}
                >
                  {Array.from({ length: benchesPerRow }).map((_, b) => (
                    <div
                      key={b}
                      className="p-2.5 rounded-xl bg-white border border-slate-200 text-center shadow-2xs"
                    >
                      <div className="text-[9px] text-slate-500 font-mono mb-1 font-bold">
                        B-{r * benchesPerRow + b + 1}
                      </div>
                      <div
                        className="grid gap-1"
                        style={{
                          gridTemplateColumns: `repeat(${studentsPerBench}, minmax(0, 1fr))`,
                        }}
                      >
                        {Array.from({ length: studentsPerBench }).map((_, p) => (
                          <div
                            key={p}
                            className="py-1 rounded bg-[#B8B5FF]/30 border border-[#B8B5FF]/60 text-[9px] font-mono text-slate-900 font-black"
                          >
                            {["F-1", "S-1", "T-1", "F-2"][p]}
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}
