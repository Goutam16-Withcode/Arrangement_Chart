"use client";
import React, { useState } from "react";
import { motion } from "framer-motion";
import { cn } from "@/lib/utils";

type Tab = {
  title: string;
  value: string;
  content?: string | React.ReactNode;
  icon?: React.ReactNode;
};

export const Tabs = ({
  tabs: propTabs,
  containerClassName,
  activeTabClassName,
  tabClassName,
  contentClassName,
  onChange,
}: {
  tabs: Tab[];
  containerClassName?: string;
  activeTabClassName?: string;
  tabClassName?: string;
  contentClassName?: string;
  onChange?: (tab: Tab) => void;
}) => {
  const [active, setActive] = useState<Tab>(propTabs[0]);

  const moveSelectedTabToTop = (idx: number) => {
    setActive(propTabs[idx]);
    if (onChange) {
      onChange(propTabs[idx]);
    }
  };

  return (
    <>
      <div
        className={cn(
          "flex flex-row items-center justify-start relative overflow-auto sm:overflow-visible no-visible-scrollbar max-w-full w-full gap-2 p-1.5 bg-slate-200/60 border border-slate-300/60 rounded-2xl",
          containerClassName
        )}
      >
        {propTabs.map((tab, idx) => (
          <button
            key={tab.title}
            onClick={() => {
              moveSelectedTabToTop(idx);
            }}
            className={cn(
              "relative px-4 py-2.5 rounded-xl text-xs font-bold transition duration-200 flex items-center gap-2",
              tabClassName
            )}
            style={{
              transformStyle: "preserve-3d",
            }}
          >
            {active.value === tab.value && (
              <motion.div
                layoutId="clickedbutton"
                transition={{ type: "spring", bounce: 0.25, duration: 0.5 }}
                className={cn(
                  "absolute inset-0 bg-[#161618] rounded-xl shadow-md",
                  activeTabClassName
                )}
              />
            )}

            <span
              className={cn(
                "relative z-20 flex items-center gap-2 transition font-bold",
                active.value === tab.value ? "text-white" : "text-slate-600 hover:text-black"
              )}
            >
              {tab.icon && (
                <span className={active.value === tab.value ? "text-[#D4F754]" : ""}>
                  {tab.icon}
                </span>
              )}
              {tab.title}
            </span>
          </button>
        ))}
      </div>
      <FadeInDiv
        tabs={propTabs}
        active={active}
        key={active.value}
        className={cn("mt-6", contentClassName)}
      />
    </>
  );
};

export const FadeInDiv = ({
  className,
  tabs,
  active,
}: {
  className?: string;
  key?: string;
  tabs: Tab[];
  active: Tab;
}) => {
  return (
    <div className="relative w-full h-full">
      <motion.div
        key={active.value}
        initial={{ opacity: 0, y: 10 }}
        animate={{ opacity: 1, y: 0 }}
        exit={{ opacity: 0, y: -10 }}
        transition={{ duration: 0.3 }}
        className={cn("w-full h-full", className)}
      >
        {active.content}
      </motion.div>
    </div>
  );
};
