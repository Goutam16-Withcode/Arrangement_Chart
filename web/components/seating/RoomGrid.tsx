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
        colors: ["#10b981", "#34d399", "#059669"],
      });

      setSelectedSeat(null);
    }
  };

  const occupancyRate = Math.round((assignedCount / totalCapacity) * 100);

  return (
    <div className="space-y-6">
      {/* Room Header Controls */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 p-5 rounded-2xl bg-white border border-[#E8E2D4] shadow-sm">
        <div>
          <div className="flex items-center gap-2.5">
            <h2 className="text-xl font-bold text-slate-900 tracking-tight">
              Room {roomConfig.roomNumber}
            </h2>
            <span className="text-xs px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-800 border border-emerald-200 font-medium">
              {roomConfig.building} • Floor {roomConfig.floor}
            </span>
          </div>
          <p className="text-xs text-slate-500 mt-1">
            Layout: {roomConfig.rows} Rows × {roomConfig.benchesPerRow} Benches ({roomConfig.studentsPerBench} students per bench)
          </p>
        </div>

        {/* Live occupancy & conflict status */}
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2 px-3.5 py-1.5 rounded-xl bg-[#F8F6F0] border border-[#E8E2D4]">
            <Users className="h-4 w-4 text-emerald-600" />
            <div className="text-xs">
              <span className="text-slate-900 font-bold">{assignedCount}</span>
              <span className="text-slate-500">/{totalCapacity} Seated</span>
              <span className="ml-1.5 text-emerald-700 font-mono font-bold">({occupancyRate}%)</span>
            </div>
          </div>

          <div className="flex items-center gap-2 px-3.5 py-1.5 rounded-xl bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-semibold">
            <ShieldCheck className="h-4 w-4 text-emerald-600" />
            <span>Anti-Cheating Active</span>
          </div>
        </div>
      </div>

      {/* Seat Swap Active Floating Notification */}
      {selectedSeat && (
        <div className="p-3.5 bg-emerald-100 border border-emerald-300 rounded-xl flex items-center justify-between shadow-sm">
          <div className="flex items-center gap-2 text-xs text-emerald-900 font-medium">
            <ArrowLeftRight className="h-4 w-4 text-emerald-700" />
            <span>
              Selected <strong>{selectedSeat.student ? selectedSeat.student.rollNo : selectedSeat.positionLabel}</strong> (Bench #{selectedSeat.benchIndex + 1}). Click any other seat to swap positions!
            </span>
          </div>
          <button
            onClick={() => setSelectedSeat(null)}
            className="text-xs px-2.5 py-1 rounded bg-white hover:bg-emerald-50 text-emerald-800 border border-emerald-300 font-semibold transition"
          >
            Cancel
          </button>
        </div>
      )}

      {/* 2D Classroom Environment View */}
      <div className="p-6 rounded-3xl bg-white border border-[#E8E2D4] shadow-sm space-y-6 relative overflow-x-auto">
        {/* Stage & Teacher Podium Blackboard */}
        <div className="w-full py-3 px-6 rounded-2xl bg-[#2A3439] text-white text-center relative shadow-sm">
          <div className="absolute left-6 top-1/2 -translate-y-1/2 flex items-center gap-2 text-[10px] text-emerald-300 uppercase tracking-widest font-mono">
            <span className="h-2 w-2 rounded-full bg-emerald-400 animate-ping" />
            Front Door
          </div>
          <span className="text-xs font-semibold text-emerald-100 tracking-wider uppercase">
            🖥️ Front Screen & Invigilator Stage / Podium
          </span>
          <div className="absolute right-6 top-1/2 -translate-y-1/2 text-[10px] text-emerald-300 uppercase tracking-widest font-mono">
            Exit 🚪
          </div>
        </div>

        {/* Matrix of Benches */}
        <div className="space-y-6 min-w-[650px]">
          {Array.from({ length: roomConfig.rows }).map((_, rowIndex) => {
            return (
              <div key={`row-${rowIndex}`} className="space-y-1.5">
                <div className="flex items-center justify-between text-[11px] font-mono text-slate-400 px-1 font-semibold">
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
                          className="p-3 rounded-2xl bg-[#FAF8F3] border border-[#E9E4D8] hover:border-emerald-300 transition shadow-2xs"
                        >
                          <div className="text-[10px] font-mono text-slate-500 mb-2 flex items-center justify-between font-semibold">
                            <span>Bench {benchIndex + 1}</span>
                            <span className="text-emerald-700">
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
