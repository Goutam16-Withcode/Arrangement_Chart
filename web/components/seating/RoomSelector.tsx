"use client";
import React from "react";
import { useSeating } from "@/lib/context/SeatingContext";
import { DoorOpen, Plus } from "lucide-react";
import Link from "next/link";

export const RoomSelector = () => {
  const { roomSeatings, selectedRoomNumber, setSelectedRoomNumber } = useSeating();

  return (
    <div className="flex items-center gap-2 overflow-x-auto pb-2 no-visible-scrollbar">
      {roomSeatings.map((rs) => {
        const isSelected = rs.roomConfig.roomNumber === selectedRoomNumber;
        const occupancyRate = Math.round(
          (rs.assignedCount / rs.totalCapacity) * 100
        );

        return (
          <button
            key={rs.roomConfig.roomNumber}
            onClick={() => setSelectedRoomNumber(rs.roomConfig.roomNumber)}
            className={`flex-shrink-0 px-4 py-2.5 rounded-2xl border text-left transition-all duration-200 flex items-center gap-3 ${
              isSelected
                ? "bg-gradient-to-r from-indigo-900/60 to-purple-900/60 border-indigo-500 text-white shadow-lg shadow-indigo-500/20"
                : "bg-slate-900/60 border-slate-800 text-slate-300 hover:border-slate-700 hover:bg-slate-800/60"
            }`}
          >
            <div
              className={`p-2 rounded-xl ${
                isSelected
                  ? "bg-indigo-500 text-white"
                  : "bg-slate-800 text-slate-400"
              }`}
            >
              <DoorOpen className="h-4 w-4" />
            </div>

            <div>
              <div className="text-xs font-bold font-mono">
                Room {rs.roomConfig.roomNumber}
              </div>
              <div className="text-[10px] text-slate-400">
                {rs.assignedCount}/{rs.totalCapacity} ({occupancyRate}%)
              </div>
            </div>
          </button>
        );
      })}

      <Link
        href="/builder"
        className="flex-shrink-0 px-3.5 py-2.5 rounded-2xl border border-dashed border-slate-800 hover:border-indigo-500 text-slate-400 hover:text-indigo-300 text-xs font-medium flex items-center gap-1.5 transition"
      >
        <Plus className="h-4 w-4" />
        <span>New Room</span>
      </Link>
    </div>
  );
};
