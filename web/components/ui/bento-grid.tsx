import { cn } from "@/lib/utils";

export const BentoGrid = ({
  className,
  children,
}: {
  className?: string;
  children?: React.ReactNode;
}) => {
  return (
    <div
      className={cn(
        "grid md:auto-rows-[18rem] grid-cols-1 md:grid-cols-3 gap-5 max-w-7xl mx-auto",
        className
      )}
    >
      {children}
    </div>
  );
};

export const BentoGridItem = ({
  className,
  title,
  description,
  header,
  icon,
}: {
  className?: string;
  title?: string | React.ReactNode;
  description?: string | React.ReactNode;
  header?: React.ReactNode;
  icon?: React.ReactNode;
}) => {
  return (
    <div
      className={cn(
        "row-span-1 rounded-2xl group/bento transition duration-300 p-5 bg-white border border-[#E8E2D4] shadow-sm hover:shadow-lg hover:shadow-emerald-900/5 justify-between flex flex-col space-y-4 hover:border-emerald-400/60",
        className
      )}
    >
      {header}
      <div className="group-hover/bento:translate-x-1 transition duration-200">
        <div className="flex items-center gap-2 mb-1.5">
          {icon}
          <div className="font-bold text-slate-800 tracking-tight text-base">
            {title}
          </div>
        </div>
        <div className="font-normal text-slate-500 text-xs leading-relaxed">
          {description}
        </div>
      </div>
    </div>
  );
};
