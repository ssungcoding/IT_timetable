import assert from "node:assert/strict";
import test from "node:test";
import {
  buildScheduleCandidates,
  createEmptyAssignments,
  createEmptyBlocked,
  createEmptyOperatingGrid,
  DAYS,
  recalculateScheduleResult,
  TIMES,
  updateStandbyAssignment,
} from "../app/scheduler.ts";

const legacyWorkSlots = [6, 7, 8, 9, 12, 13, 14, 15, 16, 17, 18, 19];
const createWeekdayOperating = () =>
  DAYS.map((_, day) => TIMES.map((__, slot) =>
    day < 5 && legacyWorkSlots.includes(slot)
  ));

for (const peopleCount of [5, 6]) {
  test(`${peopleCount}명의 대기시간을 동일하게 배정한다`, () => {
    const blocked = createEmptyBlocked(peopleCount);
    const operating = createWeekdayOperating();
    const result = buildScheduleCandidates(blocked, 1, 1, operating)[0];
    const spread = Math.max(...result.standbyHours) - Math.min(...result.standbyHours);

    assert.equal(result.unfilledStandby, 0);
    assert.equal(spread, 0);
    assert.equal(
      result.standbyHours.reduce((sum, hours) => sum + hours, 0),
      30,
    );

    for (let day = 0; day < DAYS.length; day += 1) {
      for (let slot = 0; slot < TIMES.length; slot += 1) {
        const standby = result.standbyAssignments[day][slot];
        if (operating[day][slot]) {
          assert.notEqual(standby, null);
          assert.notEqual(standby, result.assignments[day][slot]);
          assert.equal(blocked[standby!][day][slot], false);
        } else {
          assert.equal(standby, null);
          assert.equal(result.assignments[day][slot], null);
        }
      }
    }
  });
}

test("서로 다른 추천 후보를 20개 생성한다", () => {
  const results = buildScheduleCandidates(
    createEmptyBlocked(5),
    20,
    1,
    createWeekdayOperating(),
  );
  const signatures = new Set(
    results.map((result) => result.assignments.flat().join("")),
  );

  assert.equal(results.length, 20);
  assert.equal(signatures.size, 20);
});

test("사전 고정 배정은 유지하고 나머지 운영 칸만 자동 배정한다", () => {
  const blocked = createEmptyBlocked(5);
  const operating = createWeekdayOperating();
  const fixedAssignments = createEmptyAssignments();
  fixedAssignments[0][legacyWorkSlots[0]] = 2;
  fixedAssignments[2][legacyWorkSlots[5]] = 4;

  const results = buildScheduleCandidates(
    blocked,
    10,
    1,
    operating,
    fixedAssignments,
  );

  assert.equal(results.length, 10);
  results.forEach((result) => {
    assert.equal(result.assignments[0][legacyWorkSlots[0]], 2);
    assert.equal(result.assignments[2][legacyWorkSlots[5]], 4);
    assert.equal(result.assignments.flat().filter((person) => person !== null).length, 60);
  });
});

test("수업 중인 학생을 사전 고정하면 생성하지 않는다", () => {
  const blocked = createEmptyBlocked(5);
  const operating = createWeekdayOperating();
  const fixedAssignments = createEmptyAssignments();
  const slot = legacyWorkSlots[0];
  blocked[1][0][slot] = true;
  fixedAssignments[0][slot] = 1;

  assert.throws(
    () => buildScheduleCandidates(blocked, 1, 1, operating, fixedAssignments),
    /고정한 학생은 근로할 수 없습니다/,
  );
});

test("시간별 배정 제외 학생은 근로와 대기 모두에 배정하지 않는다", () => {
  const blocked = createEmptyBlocked(5);
  const excluded = createEmptyBlocked(5);
  const operating = createWeekdayOperating();
  const slot = legacyWorkSlots[0];
  excluded[0][0][slot] = true;
  excluded[1][0][slot] = true;

  const results = buildScheduleCandidates(
    blocked,
    5,
    1,
    operating,
    createEmptyAssignments(),
    excluded,
  );

  results.forEach((result) => {
    assert.notEqual(result.assignments[0][slot], 0);
    assert.notEqual(result.assignments[0][slot], 1);
    assert.notEqual(result.standbyAssignments[0][slot], 0);
    assert.notEqual(result.standbyAssignments[0][slot], 1);
  });
});

test("한 시간에 모든 학생을 배정 제외하면 생성하지 않는다", () => {
  const blocked = createEmptyBlocked(3);
  const excluded = createEmptyBlocked(3);
  const operating = createEmptyOperatingGrid();
  operating[0][0] = true;
  excluded.forEach((person) => {
    person[0][0] = true;
  });

  assert.throws(
    () => buildScheduleCandidates(
      blocked,
      1,
      1,
      operating,
      createEmptyAssignments(),
      excluded,
    ),
    /배정 제외로 근로 가능한 학생이 없는 시간/,
  );
});

test("수동 변경 후 개인별 시간과 대기표를 다시 계산한다", () => {
  const blocked = createEmptyBlocked(5);
  const operating = createWeekdayOperating();
  const original = buildScheduleCandidates(blocked, 1, 1, operating)[0];
  const assignments = original.assignments.map((day) => [...day]);
  const targetSlot = legacyWorkSlots[0];
  const previousPerson = assignments[0][targetSlot]!;
  const nextPerson = (previousPerson + 1) % 5;
  assignments[0][targetSlot] = nextPerson;

  const recalculated = recalculateScheduleResult(assignments, blocked, 1, operating);

  assert.equal(recalculated.hours[previousPerson], original.hours[previousPerson] - 0.5);
  assert.equal(recalculated.hours[nextPerson], original.hours[nextPerson] + 0.5);
  assert.equal(recalculated.hours.reduce((sum, hours) => sum + hours, 0), 30);
  assert.notEqual(recalculated.standbyAssignments[0][targetSlot], nextPerson);
});

test("최소 1일로 설정해도 필요하면 2일 출근으로 확장한다", () => {
  const blocked = createEmptyBlocked(5);
  for (let day = 0; day < DAYS.length; day += 1) {
    for (let slot = 0; slot < TIMES.length; slot += 1) {
      blocked[0][day][slot] = !(day < 2 && legacyWorkSlots.slice(0, 6).includes(slot));
    }
  }

  const result = buildScheduleCandidates(blocked, 1, 1, createWeekdayOperating())[0];

  assert.equal(result.hours[0], 6);
  assert.equal(result.attendanceDays[0], 2);
  assert.equal(result.warnings.some((warning) => warning.startsWith("1번 학생의 근로시간")), false);
});

test("최소 2일로 설정해도 근로시간 확보에 필요하면 3일로 확장한다", () => {
  const blocked = createEmptyBlocked(6);
  for (let day = 0; day < DAYS.length; day += 1) {
    for (let slot = 0; slot < TIMES.length; slot += 1) {
      blocked[0][day][slot] = !(day < 3 && legacyWorkSlots.slice(0, 3).includes(slot));
    }
  }

  const result = buildScheduleCandidates(blocked, 1, 2, createWeekdayOperating())[0];

  assert.ok(result.hours[0] >= 4 && result.hours[0] <= 6);
  assert.equal(result.attendanceDays[0], 3);
  assert.equal(result.warnings.some((warning) => warning.startsWith("1번 학생의 근로시간")), false);
});

for (const peopleCount of [3, 4, 7, 8, 9, 10]) {
  test(`${peopleCount}명도 선택한 운영시간을 공평하게 배정한다`, () => {
    const operating = createWeekdayOperating();
    const result = buildScheduleCandidates(
      createEmptyBlocked(peopleCount),
      1,
      1,
      operating,
    )[0];
    const spread = Math.max(...result.hours) - Math.min(...result.hours);

    assert.ok(spread <= 0.5);
    assert.equal(result.hours.reduce((sum, hours) => sum + hours, 0), 30);
    assert.equal(result.assignments[5].every((person) => person === null), true);
    assert.equal(result.assignments[6].every((person) => person === null), true);
  });
}

test("색칠한 요일과 시간에만 근로 및 대기를 편성한다", () => {
  const operating = createEmptyOperatingGrid();
  operating[0][10] = true;
  operating[0][11] = true;
  operating[6][0] = true;
  operating[6][1] = true;
  const result = buildScheduleCandidates(createEmptyBlocked(3), 1, 1, operating)[0];

  for (let day = 0; day < DAYS.length; day += 1) {
    for (let slot = 0; slot < TIMES.length; slot += 1) {
      if (operating[day][slot]) {
        assert.notEqual(result.assignments[day][slot], null);
      } else {
        assert.equal(result.assignments[day][slot], null);
        assert.equal(result.standbyAssignments[day][slot], null);
      }
    }
  }
  assert.equal(result.hours.reduce((sum, hours) => sum + hours, 0), 2);
  assert.equal(result.workDays.flat().includes("일"), true);
});

test("운영시간을 선택하지 않으면 생성하지 않는다", () => {
  assert.throws(
    () => buildScheduleCandidates(createEmptyBlocked(3), 1, 1, createEmptyOperatingGrid()),
    /근로를 운영할 시간을 먼저 색칠/,
  );
});

test("일요일을 포함해 최소 7일 출근도 설정할 수 있다", () => {
  const operating = DAYS.map(() => TIMES.map((_, slot) => slot < 3));
  const result = buildScheduleCandidates(createEmptyBlocked(3), 1, 7, operating)[0];

  assert.deepEqual(result.attendanceDays, [7, 7, 7]);
  assert.equal(result.workDays.every((days) => days.includes("일")), true);
});

test("07시부터 22시까지 점심시간을 포함해 선택할 수 있다", () => {
  assert.equal(TIMES[0], "07:00~07:30");
  assert.equal(TIMES.at(-1), "21:30~22:00");
  assert.equal(TIMES.includes("12:00~12:30"), true);
  assert.equal(TIMES.includes("12:30~13:00"), true);
});

test("대기시간표 수정 시 가능한 학생만 배정하고 대기시간을 갱신한다", () => {
  const operating = createEmptyOperatingGrid();
  operating[0][0] = true;
  operating[0][1] = true;
  const blocked = createEmptyBlocked(3);
  const result = buildScheduleCandidates(blocked, 1, 1, operating)[0];
  const currentStandby = result.standbyAssignments[0][0]!;
  const nextStandby = [0, 1, 2].find((person) =>
    person !== currentStandby && person !== result.assignments[0][0]
  )!;
  const updated = updateStandbyAssignment(result, blocked, operating, 0, 0, nextStandby);

  assert.equal(updated.standbyAssignments[0][0], nextStandby);
  assert.equal(updated.standbyHours[currentStandby], result.standbyHours[currentStandby] - 0.5);
  assert.equal(updated.standbyHours[nextStandby], result.standbyHours[nextStandby] + 0.5);
  assert.throws(
    () => updateStandbyAssignment(
      result,
      blocked,
      operating,
      0,
      0,
      result.assignments[0][0]!,
    ),
    /대기 근로자로 배정할 수 없습니다/,
  );
});

test("배정 제외된 학생으로 대기시간표를 수정할 수 없다", () => {
  const operating = createEmptyOperatingGrid();
  operating[0][0] = true;
  const blocked = createEmptyBlocked(3);
  const excluded = createEmptyBlocked(3);
  const result = buildScheduleCandidates(blocked, 1, 1, operating)[0];
  const excludedPerson = [0, 1, 2].find(
    (person) => person !== result.assignments[0][0],
  )!;
  excluded[excludedPerson][0][0] = true;

  assert.throws(
    () => updateStandbyAssignment(
      result,
      blocked,
      operating,
      0,
      0,
      excludedPerson,
      excluded,
    ),
    /대기 근로자로 배정할 수 없습니다/,
  );
});
