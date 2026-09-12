"use client";
import React from "react";
import Link from "next/link";
import { Spotlight } from "@/components/ui/spotlight";
import { SparklesCore } from "@/components/ui/sparkles";
import { Button as MovingBorderButton } from "@/components/ui/moving-border";
import { HoverEffect } from "@/components/ui/card-hover-effect";
import { DropZone } from "@/components/upload/DropZone";
import {
  Grid3X3,
  Layers,
  Search,
  Printer,
  ShieldCheck,
  Zap,
  ArrowRight,
  Sliders,
  CheckCircle2,
} from "lucide-react";
import { useSeating } from "@/lib/context/SeatingContext";

export default function HomePage() {
  const { metrics } = useSeating();

  const featureCards = [
    {
      title: "Interactive 2D Seating Studio",
      description:
        "Visualize classroom benches in real time. Drag, swap, and inspect student seat assignments with zero-delay updates.",
      link: "/studio",
      icon: <Grid3X3 className="h-6 w-6" />,
      badge: "Core Studio",
    },
    {
      title: "Constraint Anti-Cheating Engine",
      description:
        "Multi-branch interleaving algorithm ensures no two students with the same subject or department sit next to each other.",
      link: "/studio",
      icon: <ShieldCheck className="h-6 w-6" />,
      badge: "AI Powered",
    },
    {
      title: "2D Visual Room Blueprint Builder",
      description:
        "Design lecture halls, auditoriums, and computer labs with customizable bench sizes (1-4 seats), doors, and podiums.",
      link: "/builder",
      icon: <Layers className="h-6 w-6" />,
      badge: "Drag & Drop",
    },
    {
      title: "Student Kiosk & Digital Admit Pass",
      description:
        "Instant roll number search portal with 3D pin classroom locators and verification QR codes for exam morning.",
      link: "/find-seat",
      icon: <Search className="h-6 w-6" />,
      badge: "Self Service",
    },
    {
      title: "Print Station & Multi-Format Exporter",
      description:
        "1-click batch export for A4 door charts, photo attendance sheets with signature blocks, desk tags, and styled Excel workbooks.",
      link: "/export",
      icon: <Printer className="h-6 w-6" />,
      badge: "Export",
    },
    {
      title: "Faculty Invigilation Manager",
      description:
        "Synchronize proctors and exam coordinators directly with hall capacities and emergency relief rotations.",
      link: "/dashboard",
      icon: <Zap className="h-6 w-6" />,
      badge: "Roster",
    },
  ];

  return (
    <div className="relative overflow-hidden min-h-screen">
      {/* Background Spotlight */}
      <Spotlight
        className="-top-40 left-0 md:left-60 md:-top-20"
        fill="#818cf8"
      />

      {/* Hero Section */}
      <section className="relative pt-20 pb-16 px-4 sm:px-6 lg:px-8 max-w-7xl mx-auto z-10 text-center">
        <div className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-indigo-500/10 border border-indigo-500/20 text-indigo-300 text-xs font-medium mb-8 animate-pulse-glow">
          <Sparkles className="h-3.5 w-3.5 text-indigo-400" />
          <span>Next-Generation Algorithmic Exam Seating Studio</span>
        </div>

        <h1 className="text-4xl sm:text-6xl lg:text-7xl font-extrabold tracking-tight text-white max-w-4xl mx-auto leading-tight">
          Exam Seating without{" "}
          <span className="text-gradient">Conflicts</span>. Automated in{" "}
          <span className="text-gradient-cyan">Seconds</span>.
        </h1>

        <p className="mt-6 text-base sm:text-lg text-slate-400 max-w-2xl mx-auto leading-relaxed">
          Upgrade from static spreadsheets to a dynamic 2D visual seating suite.
          Features constraint-satisfaction branch interleaving, visual room builders,
          and digital QR attendance passes.
        </p>

        {/* CTA Buttons */}
        <div className="mt-10 flex flex-wrap items-center justify-center gap-4">
          <Link href="/studio">
            <MovingBorderButton
              borderRadius="1.25rem"
              className="bg-slate-950 text-white border-slate-800 px-6 py-3 font-semibold text-sm flex items-center gap-2 hover:bg-slate-900 transition"
            >
              <span>Launch Seating Studio</span>
              <ArrowRight className="h-4 w-4 text-indigo-400" />
            </MovingBorderButton>
          </Link>

          <Link
            href="/dashboard"
            className="px-6 py-3 rounded-2xl bg-slate-900/80 hover:bg-slate-800 text-slate-300 hover:text-white border border-slate-700/60 text-sm font-semibold transition flex items-center gap-2"
          >
            <Sliders className="h-4 w-4" />
            <span>Open Dashboard</span>
          </Link>
        </div>

        {/* Sparkles Particle Divider */}
        <div className="w-full h-24 relative mt-12">
          <div className="absolute inset-x-20 top-0 bg-gradient-to-r from-transparent via-indigo-500 to-transparent h-[2px] w-3/4 blur-sm" />
          <div className="absolute inset-x-20 top-0 bg-gradient-to-r from-transparent via-indigo-500 to-transparent h-px w-3/4" />
          <div className="absolute inset-x-60 top-0 bg-gradient-to-r from-transparent via-sky-500 to-transparent h-[5px] w-1/4 blur-sm" />
          <div className="absolute inset-x-60 top-0 bg-gradient-to-r from-transparent via-sky-500 to-transparent h-px w-1/4" />

          <SparklesCore
            background="transparent"
            minSize={0.4}
            maxSize={1.8}
            particleDensity={70}
            className="w-full h-full"
            particleColor="#a5b4fc"
          />

          <div className="absolute inset-0 w-full h-full bg-[#090D16] [mask-image:radial-gradient(350px_200px_at_top,transparent_20%,white)]"></div>
        </div>
      </section>

      {/* Live Metrics Showcase */}
      <section className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mb-16">
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          <div className="p-5 rounded-2xl glass-panel text-center">
            <div className="text-3xl font-extrabold text-white font-mono">
              {metrics.totalRooms}
            </div>
            <div className="text-xs text-slate-400 mt-1 uppercase tracking-wider">
              Exam Halls Ready
            </div>
          </div>

          <div className="p-5 rounded-2xl glass-panel text-center">
            <div className="text-3xl font-extrabold text-indigo-400 font-mono">
              {metrics.seatedStudents}
            </div>
            <div className="text-xs text-slate-400 mt-1 uppercase tracking-wider">
              Students Seated
            </div>
          </div>

          <div className="p-5 rounded-2xl glass-panel text-center">
            <div className="text-3xl font-extrabold text-emerald-400 font-mono">
              {metrics.conflictFreeRate}%
            </div>
            <div className="text-xs text-slate-400 mt-1 uppercase tracking-wider">
              Zero Conflict Rate
            </div>
          </div>

          <div className="p-5 rounded-2xl glass-panel text-center">
            <div className="text-3xl font-extrabold text-pink-400 font-mono">
              {metrics.utilizationRate}%
            </div>
            <div className="text-xs text-slate-400 mt-1 uppercase tracking-wider">
              Capacity Utilization
            </div>
          </div>
        </div>
      </section>

      {/* Quick Excel Ingestion Dropzone */}
      <section className="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8 mb-20">
        <div className="text-center mb-6">
          <h2 className="text-2xl font-bold text-white tracking-tight">
            Upload Layout & Roll Lists
          </h2>
          <p className="text-xs text-slate-400 mt-1">
            Drop your institutional spreadsheets to instantly calculate optimal seat plans.
          </p>
        </div>
        <DropZone />
      </section>

      {/* Feature Navigation Grid (CardHoverEffect) */}
      <section className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mb-24">
        <div className="text-center mb-8">
          <h2 className="text-2xl sm:text-3xl font-extrabold text-white tracking-tight">
            Complete Suite of Modern Tools
          </h2>
          <p className="text-xs text-slate-400 mt-1">
            Everything you need for seamless exam management in one unified platform.
          </p>
        </div>

        <HoverEffect items={featureCards} />
      </section>
    </div>
  );
}
