"use client";
import React, { useState } from "react";
import { RoomSeating, SeatSlot } from "@/lib/types";
import { SeatSlotCard } from "./SeatSlotCard";
import {
  Users,
  ShieldCheck,
  ArrowLeftRight,
  Maximize2,
  HelpCircle,
  Sparkles,
  AlertTriangle,
} from "lucide-react";
import confetti from "canvas-confetti";

interface RoomGridProps {
  roomSeating: RoomSeating;
  onSwapSeats: (seatIdA: string, seatIdB: string) => void;
}

export const RoomGrid: React.FC<RoomGridProps> = ({
  roomSeating,
  onSwapSeats,
}) => {
  const { roomConfig, totalCapacity, assignedCount, seats } = roomSeating;
  const [selectedSeat, setSelectedSeat] = useState<SeatSlot | null>(null);
  const [swapTarget, setSwapTarget] = useState<SeatSlot | null>(null);

  const handleSeatClick = (seat: SeatSlot) => {
    if (!selectedSeat) {
      // Pick first seat
      setSelectedSeat(seat);
    } else if (selectedSeat.id === seat.id) {
      // Deselect if clicked same
      setSelectedSeat(null);
    } else {
      // Swap!
      setSwapTarget(seat);
      onSwapSeats(selectedSeat.id, seat.id);

      // Trigger micro celebration if swap resolves conflict or is successful
      confetti({
        particleCount: 25,
        spread: 60,
        origin: { y: 0.7 },
        colors: ["#6366f1", "#ec4899", "#06b6d4"],
      });

      setSelectedSeat(null);
      setSwapTarget(null);
    }
  };

  const occupancyRate = Math.round((assignedCount / totalCapacity) * 100);

  return (
    <div className="space-y-6">
      {/* Room Header Controls */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 p-5 rounded-2xl glass-panel">
        <div>
          <div className="flex items-center gap-2.5">
            <h2 className="text-xl font-bold text-white tracking-tight">
              Room {roomConfig.roomNumber}
            </h2>
            <span className="text-xs px-2.5 py-0.5 rounded-full bg-indigo-500/20 text-indigo-300 border border-indigo-500/30">
              {roomConfig.building} • Floor {roomConfig.floor}
            </span>
          </div>
          <p className="text-xs text-slate-400 mt-1">
            Layout: {roomConfig.rows} Rows × {roomConfig.benchesPerRow} Benches ({roomConfig.studentsPerBench} students per bench)
          </p>
        </div>

        {/* Live occupancy & conflict status */}
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2 px-3 py-1.5 rounded-xl bg-slate-900/80 border border-slate-800">
            <Users className="h-4 w-4 text-indigo-400" />
            <div className="text-xs">
              <span className="text-white font-bold">{assignedCount}</span>
              <span className="text-slate-400">/{totalCapacity} Seated</span>
              <span className="ml-1.5 text-indigo-400 font-mono">({occupancyRate}%)</span>
            </div>
          </div>

          <div className="flex items-center gap-2 px-3 py-1.5 rounded-xl bg-emerald-950/40 border border-emerald-800/50 text-emerald-300 text-xs">
            <ShieldCheck className="h-4 w-4 text-emerald-400" />
            <span>Anti-Cheating Active</span>
          </div>
        </div>
      </div>

      {/* Seat Swap Active Floating Notification */}
      {selectedSeat && (
        <div className="p-3 bg-pink-950/80 border border-pink-500/60 rounded-xl flex items-center justify-between animate-pulse">
          <div className="flex items-center gap-2 text-xs text-pink-200">
            <ArrowLeftRight className="h-4 w-4 text-pink-400" />
            <span>
              Selected <strong>{selectedSeat.student ? selectedSeat.student.rollNo : selectedSeat.positionLabel}</strong> (Bench #{selectedSeat.benchIndex + 1}). Click any other seat to swap positions instantly!
            </span>
          </div>
          <button
            onClick={() => setSelectedSeat(null)}
            className="text-xs px-2 py-0.5 rounded bg-pink-900/60 hover:bg-pink-800 text-pink-200 border border-pink-700"
          >
            Cancel
          </button>
        </div>
      )}

      {/* 2D Classroom Environment View */}
      <div className="p-6 rounded-3xl glass-panel space-y-6 relative overflow-x-auto">
        {/* Stage & Teacher Podium Blackboard */}
        <div className="w-full py-2.5 px-6 rounded-2xl bg-gradient-to-r from-slate-950 via-slate-900 to-slate-950 border border-slate-800 text-center relative shadow-inner">
          <div className="absolute left-6 top-1/2 -translate-y-1/2 flex items-center gap-2 text-[10px] text-slate-500 uppercase tracking-widest font-mono">
            <span className="h-2 w-2 rounded-full bg-emerald-500 animate-ping" />
            Front Door
          </div>
          <span className="text-xs font-semibold text-slate-300 tracking-wider uppercase">
            🖥️ Front Screen & Invigilator Stage / Podium
          </span>
          <div className="absolute right-6 top-1/2 -translate-y-1/2 text-[10px] text-slate-500 uppercase tracking-widest font-mono">
            Exit 🚪
          </div>
        </div>

        {/* Matrix of Benches */}
        <div className="space-y-6 min-w-[650px]">
          {Array.from({ length: roomConfig.rows }).map((_, rowIndex) => {
            return (
              <div key={`row-${rowIndex}`} className="space-y-1.5">
                <div className="flex items-center justify-between text-[11px] font-mono text-slate-500 px-1">
                  <span>Row {rowIndex + 1}</span>
                  <span>Aisle Walkway</span>
                </div>

                {/* Benches in this row */}
                <div
                  className="grid gap-4"
                  style={{
                    gridTemplateColumns: `repeat(${roomConfig.benchesPerRow}, minmax(0, 1fr))`,
                  }}
                >
                  {Array.from({ length: roomConfig.benchesPerRow }).map(
                    (_, colIndex) => {
                      const benchIndex = rowIndex * roomConfig.benchesPerRow + colIndex;
                      const benchSeats = seats[benchIndex] || [];

                      return (
                        <div
                          key={`bench-${benchIndex}`}
                          className="p-2.5 rounded-2xl bg-slate-950/60 border border-slate-800/80 hover:border-slate-700 transition"
                        >
                          <div className="text-[10px] font-mono text-slate-400 mb-2 flex items-center justify-between">
                            <span>Bench {benchIndex + 1}</span>
                            <span className="text-slate-600">
                              {benchSeats.filter((s) => s.student).length}/
                              {roomConfig.studentsPerBench}
                            </span>
                          </div>

                          {/* Seats on this bench */}
                          <div
                            className="grid gap-1.5"
                            style={{
                              gridTemplateColumns: `repeat(${roomConfig.studentsPerBench}, minmax(0, 1fr))`,
                            }}
                          >
                            {benchSeats.map((seat) => (
                              <SeatSlotCard
                                key={seat.id}
                                seat={seat}
                                isSelected={selectedSeat?.id === seat.id}
                                onSelect={handleSeatClick}
                              />
                            ))}
                          </div>
                        </div>
                      );
                    }
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};
