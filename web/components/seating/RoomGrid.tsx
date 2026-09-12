"use client";
import React, { useState } from "react";
import { RoomSeating, SeatSlot } from "@/lib/types";
import { SeatSlotCard } from "./SeatSlotCard";
import {
  Users,
  ShieldCheck,
  ArrowLeftRight,
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

  const handleSeatClick = (seat: SeatSlot) => {
    if (!selectedSeat) {
      setSelectedSeat(seat);
    } else if (selectedSeat.id === seat.id) {
      setSelectedSeat(null);
    } else {
      onSwapSeats(selectedSeat.id, seat.id);

      confetti({
        particleCount: 25,
        spread: 60,
        origin: { y: 0.7 },
        colors: ["#D4F754", "#B8B5FF", "#161618"],
      });

      setSelectedSeat(null);
    }
  };

  const occupancyRate = Math.round((assignedCount / totalCapacity) * 100);

  return (
    <div className="space-y-6">
      {/* Room Header Controls */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 p-5 rounded-[24px] bg-white border border-slate-200/80 shadow-xs">
        <div>
          <div className="flex items-center gap-2.5">
            <h2 className="text-xl font-black text-slate-900 tracking-tight">
              Room {roomConfig.roomNumber}
            </h2>
            <span className="text-[11px] px-3 py-1 rounded-full bg-[#B8B5FF]/30 text-slate-900 font-bold border border-[#B8B5FF]/50">
              {roomConfig.building} • Floor {roomConfig.floor}
            </span>
          </div>
          <p className="text-xs text-slate-500 mt-1 font-medium">
            Layout: {roomConfig.rows} Rows × {roomConfig.benchesPerRow} Benches ({roomConfig.studentsPerBench} students per bench)
          </p>
        </div>

        {/* Live occupancy & conflict status */}
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 px-3.5 py-1.5 rounded-xl bg-slate-100 border border-slate-200">
            <Users className="h-4 w-4 text-slate-700" />
            <div className="text-xs font-semibold">
              <span className="text-slate-900 font-black">{assignedCount}</span>
              <span className="text-slate-500">/{totalCapacity}</span>
              <span className="ml-1.5 text-black font-mono font-bold bg-[#D4F754] px-1.5 py-0.5 rounded-md">({occupancyRate}%)</span>
            </div>
          </div>

          <div className="flex items-center gap-2 px-3.5 py-1.5 rounded-xl bg-[#161618] text-[#D4F754] text-xs font-black shadow-xs">
            <ShieldCheck className="h-4 w-4" />
            <span>Anti-Cheat AI Active</span>
          </div>
        </div>
      </div>

      {/* Seat Swap Active Floating Notification */}
      {selectedSeat && (
        <div className="p-3.5 bg-[#161618] text-white rounded-2xl flex items-center justify-between shadow-lg ring-2 ring-[#D4F754]/50 animate-bounce">
          <div className="flex items-center gap-2.5 text-xs font-medium">
            <div className="p-1 rounded-lg bg-[#D4F754] text-black">
              <ArrowLeftRight className="h-4 w-4" />
            </div>
            <span>
              Selected <strong className="text-[#D4F754]">{selectedSeat.student ? selectedSeat.student.rollNo : selectedSeat.positionLabel}</strong> (Bench #{selectedSeat.benchIndex + 1}). Click any other seat to swap positions!
            </span>
          </div>
          <button
            onClick={() => setSelectedSeat(null)}
            className="text-xs px-3 py-1 rounded-xl bg-white/20 hover:bg-white/30 text-white font-bold transition"
          >
            Cancel
          </button>
        </div>
      )}

      {/* 2D Classroom Environment View */}
      <div className="p-6 rounded-[28px] bg-white border border-slate-200/80 shadow-xs space-y-6 relative overflow-x-auto">
        {/* Stage & Teacher Podium Blackboard */}
        <div className="w-full py-3.5 px-6 rounded-2xl bg-[#161618] text-white text-center relative shadow-md">
          <div className="absolute left-6 top-1/2 -translate-y-1/2 flex items-center gap-2 text-[10px] text-[#D4F754] uppercase tracking-widest font-black">
            <span className="h-2 w-2 rounded-full bg-[#D4F754] animate-ping" />
            Front Door
          </div>
          <span className="text-xs font-black text-white tracking-wider uppercase flex items-center justify-center gap-2">
            <span>🖥️</span> Front Screen & Invigilator Stage / Podium
          </span>
          <div className="absolute right-6 top-1/2 -translate-y-1/2 text-[10px] text-[#D4F754] uppercase tracking-widest font-black">
            Exit 🚪
          </div>
        </div>

        {/* Matrix of Benches */}
        <div className="space-y-6 min-w-[650px]">
          {Array.from({ length: roomConfig.rows }).map((_, rowIndex) => {
            return (
              <div key={`row-${rowIndex}`} className="space-y-2">
                <div className="flex items-center justify-between text-[11px] font-mono text-slate-400 px-1 font-bold">
                  <span>Row {rowIndex + 1}</span>
                  <span className="text-[10px] text-slate-400 uppercase tracking-wider">Aisle Walkway</span>
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
                          className="p-3.5 rounded-2xl bg-[#F8F9FA] border border-slate-200 hover:border-black transition-all shadow-2xs"
                        >
                          <div className="text-[11px] font-bold text-slate-600 mb-2.5 flex items-center justify-between font-mono">
                            <span>Bench {benchIndex + 1}</span>
                            <span className="bg-white px-2 py-0.5 rounded-full border border-slate-200 text-slate-800 text-[10px] font-black">
                              {benchSeats.filter((s) => s.student).length}/
                              {roomConfig.studentsPerBench}
                            </span>
                          </div>

                          {/* Seats on this bench */}
                          <div
                            className="grid gap-2"
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
