import assert from "node:assert/strict";
import test from "node:test";
import {
  buildScheduleCandidates,
  createEmptyBlocked,
  createEmptyOperatingGrid,
  DAYS,
  recalculateScheduleResult,
  TIMES,
} from "../app/scheduler.ts";

const createWeekdayOperating = () =>
  DAYS.map((_, day) => TIMES.map(() => day < 5));

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
      5 * TIMES.length * 0.5,
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

test("수동 변경 후 개인별 시간과 대기표를 다시 계산한다", () => {
  const blocked = createEmptyBlocked(5);
  const operating = createWeekdayOperating();
  const original = buildScheduleCandidates(blocked, 1, 1, operating)[0];
  const assignments = original.assignments.map((day) => [...day]);
  const previousPerson = assignments[0][0]!;
  const nextPerson = (previousPerson + 1) % 5;
  assignments[0][0] = nextPerson;

  const recalculated = recalculateScheduleResult(assignments, blocked, 1, operating);

  assert.equal(recalculated.hours[previousPerson], original.hours[previousPerson] - 0.5);
  assert.equal(recalculated.hours[nextPerson], original.hours[nextPerson] + 0.5);
  assert.equal(recalculated.hours.reduce((sum, hours) => sum + hours, 0), 30);
  assert.notEqual(recalculated.standbyAssignments[0][0], nextPerson);
});

test("최소 1일로 설정해도 필요하면 2일 출근으로 확장한다", () => {
  const blocked = createEmptyBlocked(5);
  for (let day = 0; day < DAYS.length; day += 1) {
    for (let slot = 0; slot < TIMES.length; slot += 1) {
      blocked[0][day][slot] = !(day < 2 && slot < 6);
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
      blocked[0][day][slot] = !(day < 3 && slot < 3);
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
  const operating = DAYS.map(() => TIMES.map(() => true));
  const result = buildScheduleCandidates(createEmptyBlocked(3), 1, 7, operating)[0];

  assert.deepEqual(result.attendanceDays, [7, 7, 7]);
  assert.equal(result.workDays.every((days) => days.includes("일")), true);
});
