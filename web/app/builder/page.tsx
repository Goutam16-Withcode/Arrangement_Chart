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
  const { addCustomRoom, setSelectedRoomNumber } = useSeating();
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
      colors: ["#10b981", "#34d399", "#059669"],
    });

    router.push("/studio");
  };

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-8 bg-[#FBF9F4]">
      {/* Header */}
      <div>
        <div className="flex items-center gap-2 text-emerald-700 text-xs font-mono font-semibold uppercase tracking-wider mb-1">
          <Sparkles className="h-4 w-4 text-emerald-600" />
          <span>Room Geometry Designer</span>
        </div>
        <h1 className="text-2xl sm:text-3xl font-extrabold text-slate-900 tracking-tight">
          2D Visual Room Blueprint Builder
        </h1>
        <p className="text-xs text-slate-500 mt-1">
          Configure custom classroom, auditorium, or computer lab layouts with dynamic bench sizes.
        </p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
        {/* Controls Configuration Form */}
        <div className="lg:col-span-5 bg-white border border-[#E8E2D4] p-6 rounded-3xl shadow-sm space-y-6">
          <form onSubmit={handleSaveRoom} className="space-y-4">
            <h3 className="text-sm font-bold text-slate-900 uppercase tracking-wider font-mono flex items-center gap-2">
              <Sliders className="h-4 w-4 text-emerald-600" />
              <span>Room Specifications</span>
            </h3>

            {/* Room Number & Floor */}
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Room Number *
                </label>
                <input
                  type="text"
                  required
                  value={roomNumber}
                  onChange={(e) => setRoomNumber(e.target.value)}
                  className="w-full px-3 py-2 rounded-xl bg-[#FAF8F3] border border-[#E2DCCE] text-xs text-slate-900 focus:outline-none focus:border-emerald-500"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Floor Number
                </label>
                <input
                  type="number"
                  min="0"
                  max="12"
                  value={floor}
                  onChange={(e) => setFloor(Number(e.target.value))}
                  className="w-full px-3 py-2 rounded-xl bg-[#FAF8F3] border border-[#E2DCCE] text-xs text-slate-900 focus:outline-none focus:border-emerald-500"
                />
              </div>
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">
                Building / Block Name
              </label>
              <input
                type="text"
                value={building}
                onChange={(e) => setBuilding(e.target.value)}
                className="w-full px-3 py-2 rounded-xl bg-[#FAF8F3] border border-[#E2DCCE] text-xs text-slate-900 focus:outline-none focus:border-emerald-500"
              />
            </div>

            {/* Dimension Sliders */}
            <div className="space-y-3 pt-2">
              <div>
                <div className="flex justify-between text-xs font-semibold text-slate-700 mb-1">
                  <span>Number of Rows</span>
                  <span className="text-emerald-700 font-mono font-bold">{rows} Rows</span>
                </div>
                <input
                  type="range"
                  min="2"
                  max="12"
                  value={rows}
                  onChange={(e) => setRows(Number(e.target.value))}
                  className="w-full accent-emerald-600"
                />
              </div>

              <div>
                <div className="flex justify-between text-xs font-semibold text-slate-700 mb-1">
                  <span>Benches per Row</span>
                  <span className="text-emerald-700 font-mono font-bold">{benchesPerRow} Benches</span>
                </div>
                <input
                  type="range"
                  min="1"
                  max="8"
                  value={benchesPerRow}
                  onChange={(e) => setBenchesPerRow(Number(e.target.value))}
                  className="w-full accent-emerald-600"
                />
              </div>

              <div>
                <div className="flex justify-between text-xs font-semibold text-slate-700 mb-1">
                  <span>Students per Bench</span>
                  <span className="text-emerald-700 font-mono font-bold">
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
                          ? "bg-emerald-600 text-white border-emerald-600 shadow-sm"
                          : "bg-[#FAF8F3] text-slate-600 border-[#E2DCCE] hover:text-slate-900"
                      }`}
                    >
                      {n} {n === 1 ? "Seat" : "Seats"}
                    </button>
                  ))}
                </div>
              </div>
            </div>

            {/* Capacity Banner */}
            <div className="p-4 rounded-2xl bg-emerald-50 border border-emerald-200 flex items-center justify-between text-xs">
              <div>
                <div className="text-slate-600 font-medium">Calculated Capacity</div>
                <div className="text-lg font-bold font-mono text-emerald-900">
                  {totalCapacity} Student Desks
                </div>
              </div>
              <div className="text-right text-[11px] text-emerald-700 font-mono font-bold">
                {totalBenches} Total Benches
              </div>
            </div>

            <button
              type="submit"
              className="w-full py-3 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs tracking-wide uppercase transition shadow-md shadow-emerald-600/20 flex items-center justify-center gap-2"
            >
              <Plus className="h-4 w-4" />
              <span>Add Room & Generate Seating</span>
            </button>
          </form>
        </div>

        {/* Real-time 2D Blueprint Live Preview */}
        <div className="lg:col-span-7 bg-white border border-[#E8E2D4] p-6 rounded-3xl shadow-sm space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="text-sm font-bold text-slate-900 uppercase tracking-wider font-mono">
              Live Blueprint Preview
            </h3>
            <span className="text-xs px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-800 border border-emerald-200 font-mono font-bold">
              Room {roomNumber} • {totalCapacity} Seats
            </span>
          </div>

          {/* Blackboard Stage */}
          <div className="py-2.5 px-4 rounded-xl bg-[#2A3439] text-center text-xs font-semibold text-emerald-100 font-mono shadow-inner">
            🖥️ Front Screen / Invigilator Podium
          </div>

          {/* Rendered Live Blueprint Grid */}
          <div className="space-y-3 max-h-[440px] overflow-y-auto no-visible-scrollbar p-3 rounded-2xl bg-[#FAF8F3] border border-[#E9E4D8]">
            {Array.from({ length: rows }).map((_, r) => (
              <div key={r} className="space-y-1">
                <div className="text-[10px] text-slate-500 font-mono font-semibold">Row {r + 1}</div>
                <div
                  className="grid gap-2"
                  style={{
                    gridTemplateColumns: `repeat(${benchesPerRow}, minmax(0, 1fr))`,
                  }}
                >
                  {Array.from({ length: benchesPerRow }).map((_, b) => (
                    <div
                      key={b}
                      className="p-2 rounded-xl bg-white border border-[#E2DCCE] text-center shadow-2xs"
                    >
                      <div className="text-[9px] text-slate-500 font-mono mb-1 font-semibold">
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
                            className="py-1 rounded bg-emerald-50 border border-emerald-200 text-[9px] font-mono text-emerald-800 font-bold"
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
