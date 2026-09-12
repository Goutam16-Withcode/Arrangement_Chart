"use client";
import React, { useState } from "react";
import { useSeating } from "@/lib/context/SeatingContext";
import {
  Eye,
  Layers,
  Sparkles,
  Rotate3d,
  ShieldCheck,
  Maximize2,
  Sliders,
  Compass,
} from "lucide-react";

export default function DigitalTwinPage() {
  const { roomSeatings, selectedRoomNumber, collegeProfile, activeSession } = useSeating();
  const [viewMode, setViewMode] = useState<"isometric" | "topdown" | "podium">("isometric");
  const [showVisibilityHeatmap, setShowVisibilityHeatmap] = useState(true);

  const currentRoom =
    roomSeatings.find((rs) => rs.roomConfig.roomNumber === selectedRoomNumber) ||
    roomSeatings[0];

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <div className="flex items-center gap-2 text-xs font-mono font-bold uppercase tracking-wider mb-1 flex-wrap text-slate-800">
            <span>✦ {collegeProfile.collegeName}</span>
            <span className="text-slate-400">•</span>
            <span>{activeSession.title}</span>
            <span className="px-2 py-0.5 rounded-full bg-[#D4F754] text-black text-[10px] font-bold">
              {activeSession.date}
            </span>
          </div>
          <h1 className="text-2xl sm:text-3xl font-black text-slate-900 tracking-tight">
            3D Digital Twin Exam Hall Visualizer
          </h1>
          <p className="text-xs text-slate-500 mt-1">
            Perspective simulation of hall elevations, tiered rows, and invigilator line-of-sight coverage cones.
          </p>
        </div>

        {/* View Controls */}
        <div className="flex items-center gap-2 bg-white p-1 rounded-full border border-slate-200 shadow-2xs">
          <button
            onClick={() => setViewMode("isometric")}
            className={`px-3.5 py-1.5 rounded-full text-xs font-bold transition ${
              viewMode === "isometric"
                ? "bg-[#D4F754] text-black font-extrabold shadow-xs"
                : "text-slate-600"
            }`}
          >
            Isometric 3D
          </button>
          <button
            onClick={() => setViewMode("topdown")}
            className={`px-3.5 py-1.5 rounded-full text-xs font-bold transition ${
              viewMode === "topdown"
                ? "bg-[#D4F754] text-black font-extrabold shadow-xs"
                : "text-slate-600"
            }`}
          >
            2D Top-Down
          </button>
          <button
            onClick={() => setViewMode("podium")}
            className={`px-3.5 py-1.5 rounded-full text-xs font-bold transition ${
              viewMode === "podium"
                ? "bg-[#D4F754] text-black font-extrabold shadow-xs"
                : "text-slate-600"
            }`}
          >
            Proctor Podium
          </button>
        </div>
      </div>

      {/* Main 3D Stage Viewport */}
      {currentRoom && (
        <div className="bg-white border border-[#E8E2D4] p-8 rounded-3xl shadow-sm space-y-6 relative overflow-hidden">
          {/* Overlay badges */}
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <span className="text-sm font-bold text-slate-900">
                Hall #{currentRoom.roomConfig.roomNumber} ({currentRoom.roomConfig.building})
              </span>
              <span className="text-xs px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-800 font-mono font-bold">
                {currentRoom.assignedCount} Candidates Seated
              </span>
            </div>

            <label className="flex items-center gap-2 text-xs font-semibold text-slate-700 cursor-pointer">
              <input
                type="checkbox"
                checked={showVisibilityHeatmap}
                onChange={(e) => setShowVisibilityHeatmap(e.target.checked)}
                className="accent-emerald-600 rounded"
              />
              <span>Show Invigilator Sight Cone (Heatmap)</span>
            </label>
          </div>

          {/* 3D Isometric Viewport Container */}
          <div
            className="w-full min-h-[460px] bg-[#FAF8F3] border border-[#E8E2D4] rounded-2xl p-8 flex flex-col items-center justify-center transition-all duration-700 relative overflow-hidden"
            style={{
              perspective: viewMode === "isometric" ? "1200px" : "none",
            }}
          >
            {/* Front Stage / Screen */}
            <div
              className="w-3/4 py-3 bg-[#2A3439] text-emerald-100 text-center text-xs font-mono font-bold rounded-xl shadow-md mb-8 transition-transform duration-500"
              style={{
                transform:
                  viewMode === "isometric"
                    ? "rotateX(25deg) translateZ(30px)"
                    : "none",
              }}
            >
              🎤 INVIGILATOR PODIUM & DIGITAL CLOCK
            </div>

            {/* 3D Tiered Benches Grid */}
            <div
              className="space-y-4 w-full max-w-2xl transition-all duration-700"
              style={{
                transform:
                  viewMode === "isometric"
                    ? "rotateX(40deg) rotateZ(0deg) scale(0.92)"
                    : viewMode === "podium"
                    ? "rotateX(65deg) scale(1.05)"
                    : "none",
                transformStyle: "preserve-3d",
              }}
            >
              {Array.from({ length: currentRoom.roomConfig.rows }).map((_, rIndex) => {
                const elevation = rIndex * 12; // Tier elevation
                return (
                  <div
                    key={rIndex}
                    className="flex items-center justify-center gap-4 transition-transform duration-500"
                    style={{
                      transform:
                        viewMode === "isometric"
                          ? `translateZ(${elevation}px)`
                          : "none",
                    }}
                  >
                    <span className="text-[10px] font-mono text-slate-400 w-10 text-right">
                      Tier {rIndex + 1}
                    </span>

                    {/* Benches in this row */}
                    <div className="flex items-center gap-3">
                      {Array.from({
                        length: currentRoom.roomConfig.benchesPerRow,
                      }).map((_, bIndex) => {
                        const benchIdx =
                          rIndex * currentRoom.roomConfig.benchesPerRow + bIndex;
                        const benchSeats = currentRoom.seats[benchIdx] || [];
                        const hasStudent = benchSeats.some((s) => s.student);

                        // Line of sight rating
                        const isHighVisibility = rIndex <= 2;

                        return (
                          <div
                            key={bIndex}
                            className={`p-2.5 rounded-xl border transition-all duration-300 shadow-sm text-center flex flex-col items-center justify-center min-w-[110px] ${
                              showVisibilityHeatmap && isHighVisibility
                                ? "bg-emerald-50 border-emerald-300 text-emerald-900"
                                : showVisibilityHeatmap
                                ? "bg-amber-50/70 border-amber-200 text-slate-800"
                                : "bg-white border-[#E0D9CB] text-slate-800"
                            } hover:scale-110 hover:shadow-lg cursor-pointer`}
                          >
                            <div className="text-[9px] font-mono font-bold text-slate-500">
                              Bench #{benchIdx + 1}
                            </div>
                            <div className="text-xs font-mono font-extrabold mt-0.5">
                              {benchSeats[0]?.student
                                ? benchSeats[0].student.rollNo.slice(-6)
                                : "Vacant"}
                            </div>
                            {showVisibilityHeatmap && (
                              <div className="text-[8px] font-mono text-emerald-700 mt-1">
                                {isHighVisibility ? "● 98% Sightline" : "● 82% Sightline"}
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Visibility Legend */}
          <div className="flex items-center justify-between text-xs text-slate-500 pt-2 border-t border-[#E8E2D4] font-mono">
            <div className="flex items-center gap-4">
              <span className="flex items-center gap-1.5">
                <span className="h-3 w-3 rounded-full bg-emerald-500" />
                <span>Clear Sightline (Tier 1-3)</span>
              </span>
              <span className="flex items-center gap-1.5">
                <span className="h-3 w-3 rounded-full bg-amber-400" />
                <span>Moderate Patrol Zone (Tier 4-6)</span>
              </span>
            </div>
            <span>Camera Mode: {viewMode.toUpperCase()}</span>
          </div>
        </div>
      )}
    </div>
  );
}
