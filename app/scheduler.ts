export const DAYS = ["월", "화", "수", "목", "금"] as const;
export const TIMES = [
  "10:00~10:30",
  "10:30~11:00",
  "11:00~11:30",
  "11:30~12:00",
  "12:00~12:30",
  "12:30~13:00",
  "13:00~13:30",
  "13:30~14:00",
  "14:00~14:30",
  "14:30~15:00",
  "15:00~15:30",
  "15:30~16:00",
  "16:00~16:30",
  "16:30~17:00",
] as const;

export const isLunchSlot = (slot: number) => slot === 4 || slot === 5;

export type BlockedGrid = boolean[][][];
export type OperatingGrid = boolean[][];
export type AssignmentGrid = (number | null)[][];

export type ScheduleResult = {
  assignments: AssignmentGrid;
  standbyAssignments: (number | null)[][];
  hours: number[];
  standbyHours: number[];
  attendanceDays: number[];
  workDays: string[][];
  unfilledStandby: number;
  warnings: string[];
};

type BeamState = {
  assignment: (number | null)[];
  counts: number[];
  dayMasks: number[];
  score: number;
  last: number;
};

const popcount = (value: number) => {
  let count = 0;
  for (let n = value; n; n &= n - 1) count += 1;
  return count;
};

const transitionScore = (
  state: BeamState,
  person: number,
  day: number,
  minimumAttendanceDays: number,
) => {
  const alreadyWorkedToday = Boolean(state.dayMasks[person] & (1 << day));
  const startsNewDay = !alreadyWorkedToday;
  const switchesWorker = state.last !== -1 && state.last !== person;
  const returnsAfterGap = alreadyWorkedToday && switchesWorker;
  const existingAttendanceDays = popcount(state.dayMasks[person]);
  const newDayCost = existingAttendanceDays < minimumAttendanceDays
    ? -250
    : 80 + existingAttendanceDays * 40;

  return (
    state.score +
    (startsNewDay ? newDayCost : 0) +
    (switchesWorker ? 12 : 0) +
    (returnsAfterGap ? 45 : 0)
  );
};

const rankSchedule = (
  state: BeamState,
  peopleCount: number,
  minimumAttendanceDays: number,
  totalSlots: number,
) => {
  const target = totalSlots / peopleCount;
  const dayCounts = state.dayMasks.map(popcount);
  const totalDays = dayCounts.reduce((sum, value) => sum + value, 0);
  const daySpread = Math.max(...dayCounts) - Math.min(...dayCounts);
  const hourSpread = Math.max(...state.counts) - Math.min(...state.counts);
  const hourDeviation = state.counts.reduce(
    (sum, value) => sum + (value - target) ** 2,
    0,
  );
  const invalidHours = state.counts.reduce(
    (sum, value) => sum + Math.max(0, 8 - value) + Math.max(0, value - 12),
    0,
  );
  const minimumDayShortfall = dayCounts.reduce(
    (sum, value) => sum + Math.max(0, minimumAttendanceDays - value),
    0,
  );
  const excessAttendanceDays = dayCounts.reduce(
    (sum, value) => sum + Math.max(0, value - minimumAttendanceDays),
    0,
  );

  let switches = 0;
  for (let day = 0; day < DAYS.length; day += 1) {
    for (let slot = 1; slot < TIMES.length; slot += 1) {
      const current = state.assignment[day * TIMES.length + slot];
      const previous = state.assignment[day * TIMES.length + slot - 1];
      if (current !== null && previous !== null && current !== previous) switches += 1;
    }
  }

  return [
    minimumDayShortfall,
    invalidHours,
    hourSpread,
    hourDeviation,
    daySpread,
    excessAttendanceDays,
    totalDays,
    switches,
  ];
};

const compareRanks = (a: number[], b: number[]) => {
  for (let index = 0; index < Math.max(a.length, b.length); index += 1) {
    const difference = (a[index] ?? 0) - (b[index] ?? 0);
    if (difference !== 0) return difference;
  }
  return 0;
};

export function createEmptyBlocked(peopleCount: number): BlockedGrid {
  return Array.from({ length: peopleCount }, () =>
    Array.from({ length: DAYS.length }, () =>
      Array.from({ length: TIMES.length }, () => false),
    ),
  );
}

export function createEmptyOperatingGrid(): OperatingGrid {
  return Array.from({ length: DAYS.length }, () =>
    Array.from({ length: TIMES.length }, () => false),
  );
}

export function createEmptyAssignments(): AssignmentGrid {
  return Array.from({ length: DAYS.length }, () =>
    Array<number | null>(TIMES.length).fill(null),
  );
}

function permutations(values: number[]): number[][] {
  if (values.length <= 1) return [values];
  return values.flatMap((value, index) =>
    permutations(values.filter((_, itemIndex) => itemIndex !== index))
      .map((rest) => [value, ...rest]),
  );
}

type FlowEdge = { to: number; reverse: number; capacity: number; initial: number };
type CostFlowEdge = FlowEdge & { cost: number };

function addFlowEdge(graph: FlowEdge[][], from: number, to: number, capacity: number) {
  const forward: FlowEdge = { to, reverse: graph[to].length, capacity, initial: capacity };
  const backward: FlowEdge = { to: from, reverse: graph[from].length, capacity: 0, initial: 0 };
  graph[from].push(forward);
  graph[to].push(backward);
  return forward;
}

function calculateMaxFlow(graph: FlowEdge[][], source: number, sink: number) {
  let total = 0;
  while (true) {
    const previousNode = Array(graph.length).fill(-1) as number[];
    const previousEdge = Array(graph.length).fill(-1) as number[];
    const queue = [source];
    previousNode[source] = source;
    for (let head = 0; head < queue.length && previousNode[sink] === -1; head += 1) {
      const node = queue[head];
      graph[node].forEach((edge, edgeIndex) => {
        if (edge.capacity > 0 && previousNode[edge.to] === -1) {
          previousNode[edge.to] = node;
          previousEdge[edge.to] = edgeIndex;
          queue.push(edge.to);
        }
      });
    }
    if (previousNode[sink] === -1) break;
    let amount = Number.POSITIVE_INFINITY;
    for (let node = sink; node !== source; node = previousNode[node]) {
      amount = Math.min(amount, graph[previousNode[node]][previousEdge[node]].capacity);
    }
    for (let node = sink; node !== source; node = previousNode[node]) {
      const edge = graph[previousNode[node]][previousEdge[node]];
      edge.capacity -= amount;
      graph[node][edge.reverse].capacity += amount;
    }
    total += amount;
  }
  return total;
}

function addCostFlowEdge(
  graph: CostFlowEdge[][],
  from: number,
  to: number,
  capacity: number,
  cost: number,
) {
  const forward: CostFlowEdge = {
    to,
    reverse: graph[to].length,
    capacity,
    initial: capacity,
    cost,
  };
  const backward: CostFlowEdge = {
    to: from,
    reverse: graph[from].length,
    capacity: 0,
    initial: 0,
    cost: -cost,
  };
  graph[from].push(forward);
  graph[to].push(backward);
  return forward;
}

function calculateMinCostMaxFlow(
  graph: CostFlowEdge[][],
  source: number,
  sink: number,
) {
  let totalFlow = 0;
  while (true) {
    const distance = Array(graph.length).fill(Number.POSITIVE_INFINITY) as number[];
    const previousNode = Array(graph.length).fill(-1) as number[];
    const previousEdge = Array(graph.length).fill(-1) as number[];
    const inQueue = Array(graph.length).fill(false) as boolean[];
    const queue = [source];
    distance[source] = 0;
    inQueue[source] = true;

    for (let head = 0; head < queue.length; head += 1) {
      const node = queue[head];
      inQueue[node] = false;
      graph[node].forEach((edge, edgeIndex) => {
        const nextDistance = distance[node] + edge.cost;
        if (edge.capacity > 0 && nextDistance < distance[edge.to]) {
          distance[edge.to] = nextDistance;
          previousNode[edge.to] = node;
          previousEdge[edge.to] = edgeIndex;
          if (!inQueue[edge.to]) {
            queue.push(edge.to);
            inQueue[edge.to] = true;
          }
        }
      });
    }

    if (previousNode[sink] === -1) break;
    for (let node = sink; node !== source; node = previousNode[node]) {
      const edge = graph[previousNode[node]][previousEdge[node]];
      edge.capacity -= 1;
      graph[node][edge.reverse].capacity += 1;
    }
    totalFlow += 1;
  }
  return totalFlow;
}

function buildFairStandbyAssignments(
  assignments: AssignmentGrid,
  blocked: BlockedGrid,
  operating: OperatingGrid,
  excluded: BlockedGrid,
) {
  const peopleCount = blocked.length;
  const gridSlots = DAYS.length * TIMES.length;
  const totalSlots = operating.flat().filter(Boolean).length;
  const source = 0;
  const slotStart = 1;
  const personStart = slotStart + gridSlots;
  const sink = personStart + peopleCount;
  const graph = Array.from({ length: sink + 1 }, () => [] as CostFlowEdge[]);
  const assignmentEdges = Array.from({ length: gridSlots }, () =>
    Array<CostFlowEdge | null>(peopleCount).fill(null),
  );

  for (let index = 0; index < gridSlots; index += 1) {
    const day = Math.floor(index / TIMES.length);
    const slot = index % TIMES.length;
    if (!operating[day][slot]) continue;
    addCostFlowEdge(graph, source, slotStart + index, 1, 0);
    for (let person = 0; person < peopleCount; person += 1) {
      if (
        person !== assignments[day][slot] &&
        !blocked[person][day][slot] &&
        !excluded[person][day][slot]
      ) {
        assignmentEdges[index][person] = addCostFlowEdge(
          graph,
          slotStart + index,
          personStart + person,
          1,
          0,
        );
      }
    }
  }

  for (let person = 0; person < peopleCount; person += 1) {
    for (let count = 0; count < totalSlots; count += 1) {
      addCostFlowEdge(
        graph,
        personStart + person,
        sink,
        1,
        count * 2 + 1,
      );
    }
  }

  const filledSlots = calculateMinCostMaxFlow(graph, source, sink);
  const standbyCounts = Array(peopleCount).fill(0) as number[];
  const standbyAssignments = Array.from({ length: DAYS.length }, () =>
    Array<number | null>(TIMES.length).fill(null),
  );

  for (let index = 0; index < gridSlots; index += 1) {
    const standby = assignmentEdges[index].findIndex(
      (edge) => edge !== null && edge.initial - edge.capacity === 1,
    );
    if (standby === -1) continue;
    const day = Math.floor(index / TIMES.length);
    const slot = index % TIMES.length;
    standbyAssignments[day][slot] = standby;
    standbyCounts[standby] += 1;
  }

  return {
    standbyAssignments,
    standbyCounts,
    unfilledStandby: totalSlots - filledSlots,
  };
}

export function recalculateScheduleResult(
  assignments: AssignmentGrid,
  blocked: BlockedGrid,
  minimumAttendanceDays = 1,
  operating: OperatingGrid = createEmptyOperatingGrid(),
  excluded: BlockedGrid = createEmptyBlocked(blocked.length),
): ScheduleResult {
  operating = operating.map((day) =>
    day.map((isOperating, slot) => isOperating && !isLunchSlot(slot)),
  );
  const peopleCount = blocked.length;
  const counts = Array(peopleCount).fill(0) as number[];
  const dayMasks = Array(peopleCount).fill(0) as number[];

  for (let day = 0; day < DAYS.length; day += 1) {
    for (let slot = 0; slot < TIMES.length; slot += 1) {
      const person = assignments[day][slot];
      if (person === null || !operating[day][slot]) continue;
      counts[person] += 1;
      dayMasks[person] |= 1 << day;
    }
  }

  const attendanceDays = dayMasks.map(popcount);
  const workDays = dayMasks.map((mask) =>
    DAYS.filter((_, day) => Boolean(mask & (1 << day))),
  );
  const warnings: string[] = [];
  counts.forEach((count, person) => {
    if (count < 8 || count > 12) {
      warnings.push(`${person + 1}번 학생의 근로시간이 ${count / 2}시간입니다.`);
    }
  });
  attendanceDays.forEach((days, person) => {
    if (days < minimumAttendanceDays) {
      warnings.push(`${person + 1}번 학생의 출근 요일이 ${days}일입니다.`);
    }
  });

  const { standbyAssignments, standbyCounts, unfilledStandby } =
    buildFairStandbyAssignments(assignments, blocked, operating, excluded);

  return {
    assignments,
    standbyAssignments,
    hours: counts.map((count) => count / 2),
    standbyHours: standbyCounts.map((count) => count / 2),
    attendanceDays,
    workDays,
    unfilledStandby,
    warnings,
  };
}

export function updateStandbyAssignment(
  result: ScheduleResult,
  blocked: BlockedGrid,
  operating: OperatingGrid,
  day: number,
  slot: number,
  person: number,
  excluded: BlockedGrid = createEmptyBlocked(blocked.length),
): ScheduleResult {
  if (!operating[day]?.[slot] || isLunchSlot(slot)) {
    throw new Error("운영하지 않는 시간에는 대기 근로자를 배정할 수 없습니다.");
  }
  if (
    result.assignments[day][slot] === person ||
    blocked[person]?.[day]?.[slot] ||
    excluded[person]?.[day]?.[slot]
  ) {
    throw new Error("해당 학생은 이 시간의 대기 근로자로 배정할 수 없습니다.");
  }

  const standbyAssignments = result.standbyAssignments.map((dayAssignments) =>
    [...dayAssignments]
  );
  standbyAssignments[day][slot] = person;
  const standbyCounts = Array(blocked.length).fill(0) as number[];
  let unfilledStandby = 0;
  for (let dayIndex = 0; dayIndex < DAYS.length; dayIndex += 1) {
    for (let slotIndex = 0; slotIndex < TIMES.length; slotIndex += 1) {
      if (!operating[dayIndex][slotIndex]) continue;
      const standbyPerson = standbyAssignments[dayIndex][slotIndex];
      if (standbyPerson === null) unfilledStandby += 1;
      else standbyCounts[standbyPerson] += 1;
    }
  }

  return {
    ...result,
    standbyAssignments,
    standbyHours: standbyCounts.map((count) => count / 2),
    unfilledStandby,
  };
}

function buildExactAttendanceState(
  order: number[],
  dayShift: number,
  minimumAttendanceDays: number,
  targetSlots: number,
  blocked: BlockedGrid,
): BeamState | null {
  const peopleCount = order.length;
  const active = Array.from({ length: peopleCount }, () => Array(DAYS.length).fill(false));
  order.forEach((person, position) => {
    for (let offset = 0; offset < minimumAttendanceDays; offset += 1) {
      active[person][(position + dayShift + offset) % DAYS.length] = true;
    }
  });

  const activePerDay = DAYS.map((_, day) => active.filter((person) => person[day]).length);
  if (activePerDay.some((count) => count === 0)) return null;

  const source = 0;
  const personStart = 1;
  const dayStart = personStart + peopleCount;
  const sink = dayStart + DAYS.length;
  const graph = Array.from({ length: sink + 1 }, () => [] as FlowEdge[]);
  const allocationEdges = Array.from({ length: peopleCount }, () =>
    Array<FlowEdge | null>(DAYS.length).fill(null),
  );
  let requiredFlow = 0;

  for (let person = 0; person < peopleCount; person += 1) {
    const remaining = targetSlots - minimumAttendanceDays;
    if (remaining < 0) return null;
    addFlowEdge(graph, source, personStart + person, remaining);
    requiredFlow += remaining;
    for (let day = 0; day < DAYS.length; day += 1) {
      if (active[person][day]) {
        allocationEdges[person][day] = addFlowEdge(
          graph,
          personStart + person,
          dayStart + day,
          TIMES.length,
        );
      }
    }
  }
  for (let day = 0; day < DAYS.length; day += 1) {
    const remaining = TIMES.length - activePerDay[day];
    if (remaining < 0) return null;
    addFlowEdge(graph, dayStart + day, sink, remaining);
  }
  if (calculateMaxFlow(graph, source, sink) !== requiredFlow) return null;

  const allocations = Array.from({ length: peopleCount }, () => Array(DAYS.length).fill(0));
  for (let person = 0; person < peopleCount; person += 1) {
    for (let day = 0; day < DAYS.length; day += 1) {
      const edge = allocationEdges[person][day];
      if (edge) allocations[person][day] = 1 + edge.initial - edge.capacity;
    }
  }

  const assignment: number[] = [];
  for (let day = 0; day < DAYS.length; day += 1) {
    const daySource = 0;
    const dayPersonStart = 1;
    const slotStart = dayPersonStart + peopleCount;
    const daySink = slotStart + TIMES.length;
    const dayGraph = Array.from({ length: daySink + 1 }, () => [] as FlowEdge[]);
    const slotEdges = Array.from({ length: peopleCount }, () =>
      Array<FlowEdge | null>(TIMES.length).fill(null),
    );

    for (let person = 0; person < peopleCount; person += 1) {
      addFlowEdge(dayGraph, daySource, dayPersonStart + person, allocations[person][day]);
      for (let slot = 0; slot < TIMES.length; slot += 1) {
        if (allocations[person][day] > 0 && !blocked[person][day][slot]) {
          slotEdges[person][slot] = addFlowEdge(
            dayGraph,
            dayPersonStart + person,
            slotStart + slot,
            1,
          );
        }
      }
    }
    for (let slot = 0; slot < TIMES.length; slot += 1) {
      addFlowEdge(dayGraph, slotStart + slot, daySink, 1);
    }
    if (calculateMaxFlow(dayGraph, daySource, daySink) !== TIMES.length) return null;

    for (let slot = 0; slot < TIMES.length; slot += 1) {
      const person = slotEdges.findIndex((edges) => {
        const edge = edges[slot];
        return edge !== null && edge.initial - edge.capacity === 1;
      });
      if (person === -1) return null;
      assignment.push(person);
    }
  }

  return {
    assignment,
    counts: Array(peopleCount).fill(targetSlots),
    dayMasks: active.map((days) => days.reduce(
      (mask, isActive, day) => mask | (isActive ? 1 << day : 0),
      0,
    )),
    score: 0,
    last: -1,
  };
}

export function buildScheduleCandidates(
  blocked: BlockedGrid,
  candidateCount = 20,
  minimumAttendanceDays = 1,
  operating: OperatingGrid = createEmptyOperatingGrid(),
  fixedAssignments: AssignmentGrid = createEmptyAssignments(),
  excluded: BlockedGrid = createEmptyBlocked(blocked.length),
  preferExactHours = true,
  maxAttendanceDays?: number | null,
): ScheduleResult[] {
  operating = operating.map((day) =>
    day.map((isOperating, slot) => isOperating && !isLunchSlot(slot)),
  );
  const peopleCount = blocked.length;
  const activeSlotCount = operating.flat().filter(Boolean).length;
  if (activeSlotCount === 0) {
    throw new Error("근로를 운영할 시간을 먼저 색칠해 주세요.");
  }
  const target = activeSlotCount / peopleCount;
  const baseMaxSlots = preferExactHours
    ? Math.ceil(target)
    : Math.max(12, Math.ceil(target));
  const fixedCounts = Array(peopleCount).fill(0) as number[];
  let hasFixedAssignments = false;
  const hasExclusions = excluded.some((person) =>
    person.some((day) => day.some(Boolean)),
  );
  for (let day = 0; day < DAYS.length; day += 1) {
    for (let slot = 0; slot < TIMES.length; slot += 1) {
      const person = fixedAssignments[day]?.[slot] ?? null;
      if (person === null) continue;
      hasFixedAssignments = true;
      if (
        !Number.isInteger(person) ||
        person < 0 ||
        person >= peopleCount ||
        !operating[day][slot] ||
        blocked[person][day][slot] ||
        excluded[person][day][slot]
      ) {
        throw new Error(`${DAYS[day]}요일 ${TIMES[slot]}에 고정한 학생은 근로할 수 없습니다.`);
      }
      fixedCounts[person] += 1;
    }
  }
  const maxSlotsByPerson = fixedCounts.map((count) => Math.max(baseMaxSlots, count));
  const requiredSlots = preferExactHours
    ? Math.floor(target)
    : Math.min(8, Math.floor(target));
  const attendanceCap = maxAttendanceDays ?? null;
  const uncovered: string[] = [];

  for (let day = 0; day < DAYS.length; day += 1) {
    for (let slot = 0; slot < TIMES.length; slot += 1) {
      if (!operating[day][slot]) continue;
      const hasWorker = blocked.some((person, personIndex) =>
        !person[day][slot] && !excluded[personIndex][day][slot],
      );
      if (!hasWorker) uncovered.push(`${DAYS[day]} ${TIMES[slot]}`);
    }
  }

  if (uncovered.length > 0) {
    throw new Error(
      `수업 또는 배정 제외로 근로 가능한 학생이 없는 시간이 있습니다: ${uncovered.slice(0, 4).join(", ")}${uncovered.length > 4 ? ` 외 ${uncovered.length - 4}개` : ""}`,
    );
  }

  blocked.forEach((personGrid, person) => {
    const availableDays = personGrid.filter((dayGrid, day) =>
      dayGrid.some((cell, slot) =>
        operating[day][slot] &&
        !cell &&
        !excluded[person][day][slot] &&
        (fixedAssignments[day][slot] === null || fixedAssignments[day][slot] === person),
      ),
    ).length;
    if (availableDays < minimumAttendanceDays) {
      throw new Error(
        `${person + 1}번 학생은 근로 가능한 요일이 ${availableDays}일뿐이라 최소 ${minimumAttendanceDays}일 출근 조건을 맞출 수 없습니다.`,
      );
    }
  });

  const totalSlots = DAYS.length * TIMES.length;
  const futureAvailability = Array.from({ length: totalSlots + 1 }, () =>
    Array(peopleCount).fill(0) as number[],
  );
  for (let index = totalSlots - 1; index >= 0; index -= 1) {
    const day = Math.floor(index / TIMES.length);
    const slot = index % TIMES.length;
    for (let person = 0; person < peopleCount; person += 1) {
      futureAvailability[index][person] =
        futureAvailability[index + 1][person] +
        (
          operating[day][slot] &&
          !blocked[person][day][slot] &&
          !excluded[person][day][slot] &&
          (fixedAssignments[day][slot] === null || fixedAssignments[day][slot] === person)
            ? 1
            : 0
        );
    }
  }

  let beam: BeamState[] = [
    {
      assignment: [],
      counts: Array(peopleCount).fill(0),
      dayMasks: Array(peopleCount).fill(0),
      score: 0,
      last: -1,
    },
  ];

  for (let index = 0; index < DAYS.length * TIMES.length; index += 1) {
    const day = Math.floor(index / TIMES.length);
    const slot = index % TIMES.length;
    if (!operating[day][slot]) {
      beam = beam.map((state) => ({
        ...state,
        assignment: [...state.assignment, null],
        last: -1,
      }));
      continue;
    }
    const nextBySignature = new Map<string, BeamState>();
    const fixedPerson = fixedAssignments[day][slot];

    for (const state of beam) {
      for (let person = 0; person < peopleCount; person += 1) {
        if (
          (fixedPerson !== null && person !== fixedPerson) ||
          blocked[person][day][slot] ||
          excluded[person][day][slot] ||
          state.counts[person] >= maxSlotsByPerson[person]
        ) continue;
        const startsNewAttendanceDay = !(state.dayMasks[person] & (1 << day));
        if (
          startsNewAttendanceDay &&
          attendanceCap !== null &&
          popcount(state.dayMasks[person]) >= attendanceCap
        ) continue;

        const counts = [...state.counts];
        counts[person] += 1;
        if (counts.some(
          (count, personIndex) =>
            count + futureAvailability[index + 1][personIndex] < requiredSlots,
        )) {
          continue;
        }
        const dayMasks = [...state.dayMasks];
        dayMasks[person] |= 1 << day;
        const score = transitionScore(state, person, day, minimumAttendanceDays);
        const candidate: BeamState = {
          assignment: [...state.assignment, person],
          counts,
          dayMasks,
          score,
          last: person,
        };
        const signature = `${counts.join(",")}|${dayMasks.join(",")}|${person}`;
        const previous = nextBySignature.get(signature);
        if (!previous || candidate.score < previous.score) {
          nextBySignature.set(signature, candidate);
        }
      }
    }

    if (nextBySignature.size === 0) {
      if (attendanceCap !== null) {
        return buildScheduleCandidates(
          blocked,
          candidateCount,
          minimumAttendanceDays,
          operating,
          fixedAssignments,
          excluded,
          preferExactHours,
          null,
        );
      }
      if (preferExactHours) {
        return buildScheduleCandidates(
          blocked,
          candidateCount,
          minimumAttendanceDays,
          operating,
          fixedAssignments,
          excluded,
          false,
          null,
        );
      }
      throw new Error("입력한 수업시간으로는 모든 시간대를 채울 수 없습니다.");
    }

    beam = [...nextBySignature.values()]
      .sort((a, b) => a.score - b.score)
      .slice(0, 3500);

    if (slot === TIMES.length - 1) {
      beam = beam.map((state) => ({ ...state, last: -1 }));
    }
  }

  const constructiveStates: BeamState[] = [];
  const isFullOperating = operating.every((day) => day.every(Boolean));
  if (
    !hasFixedAssignments &&
    !hasExclusions &&
    preferExactHours &&
    isFullOperating &&
    Number.isInteger(target) &&
    peopleCount <= 6
  ) {
    const peopleOrders = permutations(Array.from({ length: peopleCount }, (_, person) => person));
    const minimumPeoplePerDay = Math.ceil(TIMES.length / target);
    if (peopleCount * minimumAttendanceDays >= DAYS.length * minimumPeoplePerDay) {
      exactCandidates:
      for (const order of peopleOrders) {
        for (let shift = 0; shift < DAYS.length; shift += 1) {
          const exactState = buildExactAttendanceState(
            order,
            shift,
            minimumAttendanceDays,
            target,
            blocked,
          );
          if (exactState) constructiveStates.push(exactState);
          if (constructiveStates.length >= candidateCount * 6) break exactCandidates;
        }
      }
    }

    for (const order of peopleOrders) {
      const counts = Array(peopleCount).fill(0) as number[];
      const dayMasks = Array(peopleCount).fill(0) as number[];
      const assignment: number[] = [];
      let last = -1;
      let valid = true;

      for (let index = 0; index < totalSlots; index += 1) {
        const day = Math.floor(index / TIMES.length);
        const slot = index % TIMES.length;
        if (slot === 0) last = -1;
        const available = order.filter(
          (person) => !blocked[person][day][slot] && counts[person] < target,
        );
        if (available.length === 0) {
          valid = false;
          break;
        }
        const person = available.includes(last) ? last : available[0];
        assignment.push(person);
        counts[person] += 1;
        dayMasks[person] |= 1 << day;
        last = person;
      }

      if (valid && counts.every((count) => count === target)) {
        constructiveStates.push({ assignment, counts, dayMasks, score: 0, last: -1 });
      }

      const balancedCounts = Array(peopleCount).fill(0) as number[];
      const balancedDayMasks = Array(peopleCount).fill(0) as number[];
      const balancedAssignment: number[] = [];
      let balancedValid = true;

      for (let index = 0; index < totalSlots; index += 1) {
        const day = Math.floor(index / TIMES.length);
        const slot = index % TIMES.length;
        let selected: number | null = null;
        for (let offset = 0; offset < peopleCount; offset += 1) {
          const person = order[(index + offset) % peopleCount];
          if (!blocked[person][day][slot] && balancedCounts[person] < target) {
            selected = person;
            break;
          }
        }
        if (selected === null) {
          balancedValid = false;
          break;
        }
        balancedAssignment.push(selected);
        balancedCounts[selected] += 1;
        balancedDayMasks[selected] |= 1 << day;
      }

      if (balancedValid && balancedCounts.every((count) => count === target)) {
        constructiveStates.push({
          assignment: balancedAssignment,
          counts: balancedCounts,
          dayMasks: balancedDayMasks,
          score: 0,
          last: -1,
        });
      }
    }
  }

  const ranked = [...constructiveStates, ...beam]
    .sort((a, b) => compareRanks(
      rankSchedule(a, peopleCount, minimumAttendanceDays, activeSlotCount),
      rankSchedule(b, peopleCount, minimumAttendanceDays, activeSlotCount),
    ))
    .filter((state, index, states) => {
      const signature = state.assignment.join("");
      return states.findIndex((item) => item.assignment.join("") === signature) === index;
    })
    .slice(0, candidateCount);

  return ranked.map((candidate) => {
    const assignments = Array.from({ length: DAYS.length }, (_, day) =>
      Array.from(
        { length: TIMES.length },
        (_, slot) => candidate.assignment[day * TIMES.length + slot],
      ),
    );
    return recalculateScheduleResult(
      assignments,
      blocked,
      minimumAttendanceDays,
      operating,
      excluded,
    );
  });
}

export function buildSchedule(
  blocked: BlockedGrid,
  operating: OperatingGrid,
): ScheduleResult {
  return buildScheduleCandidates(blocked, 1, 1, operating)[0];
}
