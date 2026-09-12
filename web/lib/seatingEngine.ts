import { RoomConfig, Student, SeatSlot, RoomSeating, SeatingMetrics } from "./types";

export const POSITION_LABELS = ["F-1", "S-1", "T-1", "F-2", "S-2"];

export class SeatingEngine {
  /**
   * Generates seating arrangement for a list of rooms and students
   * using Constraint-Satisfaction Branch Interleaving.
   */
  public static generateSeating(
    rooms: RoomConfig[],
    students: Student[],
    options: {
      enableAntiCheating?: boolean;
      prioritizeSpecialNeeds?: boolean;
    } = { enableAntiCheating: true, prioritizeSpecialNeeds: true }
  ): {
    roomSeatings: RoomSeating[];
    metrics: SeatingMetrics;
    unassignedStudents: Student[];
  } {
    // 1. Group students by branch / subject
    const branchQueues: { [branch: string]: Student[] } = {};
    const specialNeedStudents: Student[] = [];

    students.forEach((s) => {
      if (options.prioritizeSpecialNeeds && s.hasSpecialNeed) {
        specialNeedStudents.push(s);
      } else {
        if (!branchQueues[s.branch]) {
          branchQueues[s.branch] = [];
        }
        branchQueues[s.branch].push(s);
      }
    });

    const branchKeys = Object.keys(branchQueues);
    const roomSeatings: RoomSeating[] = [];
    let totalAssigned = 0;
    let totalCapacity = 0;
    const branchesCount: { [branch: string]: number } = {};

    students.forEach((s) => {
      branchesCount[s.branch] = (branchesCount[s.branch] || 0) + 1;
    });

    // 2. Process each room sequentially
    rooms.forEach((room) => {
      const totalBenches = room.rows * room.benchesPerRow;
      const capacity = totalBenches * room.studentsPerBench;
      totalCapacity += capacity;

      const seats: SeatSlot[][] = [];
      const attendanceList: {
        serialNo: number;
        position: string;
        student: Student;
        present: boolean;
      }[] = [];

      let serialCounter = 1;

      // Initialize empty seat matrix [benchIndex][positionIndex]
      for (let b = 0; b < totalBenches; b++) {
        const benchSeats: SeatSlot[] = [];
        const rowIndex = Math.floor(b / room.benchesPerRow);
        const colIndex = b % room.benchesPerRow;

        for (let p = 0; p < room.studentsPerBench; p++) {
          benchSeats.push({
            id: `seat-${room.roomNumber}-b${b}-p${p}`,
            roomNumber: room.roomNumber,
            benchIndex: b,
            positionIndex: p,
            positionLabel: POSITION_LABELS[p] || `P-${p + 1}`,
            rowIndex,
            colIndex,
            student: null,
            hasConflict: false,
          });
        }
        seats.push(benchSeats);
      }

      // Priority: Assign Special Needs students to front rows (row 0, aisle seats)
      specialNeedStudents.forEach((spStudent) => {
        for (let b = 0; b < totalBenches; b++) {
          for (let p = 0; p < room.studentsPerBench; p++) {
            if (!seats[b][p].student && seats[b][p].rowIndex === 0) {
              seats[b][p].student = spStudent;
              attendanceList.push({
                serialNo: serialCounter++,
                position: seats[b][p].positionLabel,
                student: spStudent,
                present: false,
              });
              totalAssigned++;
              return;
            }
          }
        }
      });

      // Interleaved allocation per seat column
      for (let p = 0; p < room.studentsPerBench; p++) {
        // Pick a distinct branch for this column position
        const primaryBranch = branchKeys[p % branchKeys.length];

        for (let b = 0; b < totalBenches; b++) {
          if (seats[b][p].student) continue; // already occupied by special needs

          let assignedStudent: Student | null = null;

          // 1. Try to pop from the column's assigned branch queue
          if (primaryBranch && branchQueues[primaryBranch]?.length > 0) {
            assignedStudent = branchQueues[primaryBranch].shift()!;
          } else {
            // 2. Fallback: find any non-empty branch queue
            for (const bk of branchKeys) {
              if (branchQueues[bk].length > 0) {
                assignedStudent = branchQueues[bk].shift()!;
                break;
              }
            }
          }

          if (assignedStudent) {
            seats[b][p].student = assignedStudent;
            attendanceList.push({
              serialNo: serialCounter++,
              position: seats[b][p].positionLabel,
              student: assignedStudent,
              present: false,
            });
            totalAssigned++;
          }
        }
      }

      // 3. Proximity conflict validation
      this.validateConflicts(seats, room);

      roomSeatings.push({
        roomConfig: room,
        totalCapacity: capacity,
        assignedCount: attendanceList.length,
        seats,
        attendanceList,
      });
    });

    // Unassigned students
    const unassignedStudents: Student[] = [];
    branchKeys.forEach((bk) => {
      unassignedStudents.push(...branchQueues[bk]);
    });

    // Compute conflict metrics
    let totalConflicts = 0;
    roomSeatings.forEach((rs) => {
      rs.seats.forEach((bench) => {
        bench.forEach((seat) => {
          if (seat.hasConflict) totalConflicts++;
        });
      });
    });

    const conflictFreeRate =
      totalAssigned > 0
        ? Math.max(0, 100 - (totalConflicts / totalAssigned) * 100)
        : 100;

    const metrics: SeatingMetrics = {
      totalRooms: rooms.length,
      totalCapacity,
      totalStudents: students.length,
      seatedStudents: totalAssigned,
      utilizationRate:
        totalCapacity > 0
          ? Math.round((totalAssigned / totalCapacity) * 1000) / 10
          : 0,
      conflictFreeRate: Math.round(conflictFreeRate * 10) / 10,
      branchesCount,
    };

    return { roomSeatings, metrics, unassignedStudents };
  }

  /**
   * Checks adjacency conflicts: No two students of the same branch / subject
   * should sit directly next to each other on the same bench.
   */
  public static validateConflicts(seats: SeatSlot[][], room: RoomConfig): void {
    seats.forEach((bench) => {
      for (let p = 0; p < bench.length; p++) {
        const current = bench[p];
        current.hasConflict = false;
        current.conflictReason = undefined;

        if (!current.student) continue;

        // Check horizontal neighbour on same bench
        if (p > 0 && bench[p - 1].student) {
          const prev = bench[p - 1];
          if (
            prev.student &&
            (prev.student.branch === current.student.branch ||
              prev.student.subjectCode === current.student.subjectCode)
          ) {
            current.hasConflict = true;
            current.conflictReason = `Same subject/branch with adjacent seat (${prev.positionLabel})`;
            prev.hasConflict = true;
            prev.conflictReason = `Same subject/branch with adjacent seat (${current.positionLabel})`;
          }
        }
      }
    });
  }

  /**
   * Manually swap two seats and revalidate conflicts
   */
  public static swapSeats(
    seats: SeatSlot[][],
    room: RoomConfig,
    seatIdA: string,
    seatIdB: string
  ): SeatSlot[][] {
    let slotA: SeatSlot | null = null;
    let slotB: SeatSlot | null = null;

    seats.forEach((bench) => {
      bench.forEach((seat) => {
        if (seat.id === seatIdA) slotA = seat;
        if (seat.id === seatIdB) slotB = seat;
      });
    });

    if (slotA && slotB) {
      const temp = (slotA as SeatSlot).student;
      (slotA as SeatSlot).student = (slotB as SeatSlot).student;
      (slotB as SeatSlot).student = temp;

      this.validateConflicts(seats, room);
    }

    return [...seats];
  }
}
