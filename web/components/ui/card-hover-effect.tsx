"use client";
import { cn } from "@/lib/utils";
import { AnimatePresence, motion } from "framer-motion";
import Link from "next/link";
import { useState } from "react";

export const HoverEffect = ({
  items,
  className,
}: {
  items: {
    title: string;
    description: string;
    link: string;
    icon?: React.ReactNode;
    badge?: string;
  }[];
  className?: string;
}) => {
  const [hoveredIndex, setHoveredIndex] = useState<number | null>(null);

  return (
    <div
      className={cn(
        "grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 py-6 gap-2",
        className
      )}
    >
      {items.map((item, idx) => (
        <Link
          href={item?.link}
          key={item?.link}
          className="relative group block p-2 h-full w-full"
          onMouseEnter={() => setHoveredIndex(idx)}
          onMouseLeave={() => setHoveredIndex(null)}
        >
          <AnimatePresence>
            {hoveredIndex === idx && (
              <motion.span
                className="absolute inset-0 h-full w-full bg-emerald-100/60 block rounded-3xl"
                layoutId="hoverBackground"
                initial={{ opacity: 0 }}
                animate={{
                  opacity: 1,
                  transition: { duration: 0.15 },
                }}
                exit={{
                  opacity: 0,
                  transition: { duration: 0.15, delay: 0.2 },
                }}
              />
            )}
          </AnimatePresence>
          <div className="rounded-2xl h-full w-full p-6 overflow-hidden bg-white border border-[#E8E2D4] group-hover:border-emerald-400/80 relative z-20 shadow-sm transition duration-300">
            <div className="relative z-50">
              <div className="flex items-center justify-between mb-4">
                <div className="p-3 bg-emerald-50 border border-emerald-200/60 rounded-xl text-emerald-700">
                  {item.icon}
                </div>
                {item.badge && (
                  <span className="text-[10px] font-bold uppercase tracking-wider px-2.5 py-1 bg-emerald-100 text-emerald-800 border border-emerald-200 rounded-full">
                    {item.badge}
                  </span>
                )}
              </div>
              <h4 className="text-slate-900 font-bold tracking-tight text-lg mt-2">
                {item.title}
              </h4>
              <p className="mt-2 text-slate-500 tracking-normal leading-relaxed text-xs">
                {item.description}
              </p>
            </div>
          </div>
        </Link>
      ))}
    </div>
  );
};
