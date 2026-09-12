"use client";
import React, { useState } from "react";
import {
  motion,
  AnimatePresence,
  useScroll,
  useMotionValueEvent,
} from "framer-motion";
import { cn } from "@/lib/utils";
import Link from "next/link";
import { usePathname } from "next/navigation";

export const FloatingNav = ({
  navItems,
  className,
  children,
}: {
  navItems: {
    name: string;
    link: string;
    icon?: React.ReactNode;
  }[];
  className?: string;
  children?: React.ReactNode;
}) => {
  const pathname = usePathname();
  const [hoveredIdx, setHoveredIdx] = useState<number | null>(null);

  return (
    <div
      className={cn(
        "flex max-w-fit fixed top-4 inset-x-0 mx-auto border border-[#E6E0D2] rounded-full bg-white/90 backdrop-blur-xl shadow-lg shadow-emerald-950/5 z-[5000] px-4 py-2 items-center justify-center space-x-2",
        className
      )}
    >
      {navItems.map((navItem, idx) => {
        const isActive = pathname === navItem.link;
        return (
          <Link
            key={`link-${idx}`}
            href={navItem.link}
            onMouseEnter={() => setHoveredIdx(idx)}
            onMouseLeave={() => setHoveredIdx(null)}
            className={cn(
              "relative text-xs font-semibold px-3 py-1.5 rounded-full transition duration-200 flex items-center gap-1.5",
              isActive ? "text-emerald-950 font-bold" : "text-slate-600 hover:text-emerald-800"
            )}
          >
            {/* Animated Active Pill Indicator */}
            {isActive && (
              <motion.span
                layoutId="activeNavPill"
                transition={{ type: "spring", stiffness: 380, damping: 30 }}
                className="absolute inset-0 bg-emerald-100/80 border border-emerald-300/80 rounded-full z-0"
              />
            )}

            {/* Hover ghost highlight */}
            <AnimatePresence>
              {hoveredIdx === idx && !isActive && (
                <motion.span
                  layoutId="hoverNavGhost"
                  initial={{ opacity: 0, scale: 0.95 }}
                  animate={{ opacity: 1, scale: 1 }}
                  exit={{ opacity: 0, scale: 0.95 }}
                  className="absolute inset-0 bg-[#F5F2EA] rounded-full z-0"
                />
              )}
            </AnimatePresence>

            <span className="relative z-10 flex items-center gap-1.5">
              {navItem.icon}
              <span>{navItem.name}</span>
            </span>
          </Link>
        );
      })}

      {children && <div className="relative z-10 flex items-center gap-2 pl-2 border-l border-[#E6E0D2]">{children}</div>}
    </div>
  );
};
