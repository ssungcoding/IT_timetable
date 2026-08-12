import assert from "node:assert/strict";
import test from "node:test";
import {
  buildScheduleCandidates,
  createEmptyBlocked,
  DAYS,
  recalculateScheduleResult,
  TIMES,
} from "../app/scheduler.ts";

for (const peopleCount of [5, 6]) {
  test(`${peopleCount}명의 대기시간을 동일하게 배정한다`, () => {
    const blocked = createEmptyBlocked(peopleCount);
    const result = buildScheduleCandidates(blocked, 1, 1)[0];
    const spread = Math.max(...result.standbyHours) - Math.min(...result.standbyHours);

    assert.equal(result.unfilledStandby, 0);
    assert.equal(spread, 0);
    assert.equal(
      result.standbyHours.reduce((sum, hours) => sum + hours, 0),
      DAYS.length * TIMES.length * 0.5,
    );

    for (let day = 0; day < DAYS.length; day += 1) {
      for (let slot = 0; slot < TIMES.length; slot += 1) {
        const standby = result.standbyAssignments[day][slot];
        assert.notEqual(standby, null);
        assert.notEqual(standby, result.assignments[day][slot]);
        assert.equal(blocked[standby!][day][slot], false);
      }
    }
  });
}

test("서로 다른 추천 후보를 20개 생성한다", () => {
  const results = buildScheduleCandidates(createEmptyBlocked(5), 20, 1);
  const signatures = new Set(
    results.map((result) => result.assignments.flat().join("")),
  );

  assert.equal(results.length, 20);
  assert.equal(signatures.size, 20);
});

test("수동 변경 후 개인별 시간과 대기표를 다시 계산한다", () => {
  const blocked = createEmptyBlocked(5);
  const original = buildScheduleCandidates(blocked, 1, 1)[0];
  const assignments = original.assignments.map((day) => [...day]);
  const previousPerson = assignments[0][0];
  const nextPerson = (previousPerson + 1) % 5;
  assignments[0][0] = nextPerson;

  const recalculated = recalculateScheduleResult(assignments, blocked, 1);

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

  const result = buildScheduleCandidates(blocked, 1, 1)[0];

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

  const result = buildScheduleCandidates(blocked, 1, 2)[0];

  assert.ok(result.hours[0] >= 4 && result.hours[0] <= 6);
  assert.equal(result.attendanceDays[0], 3);
  assert.equal(result.warnings.some((warning) => warning.startsWith("1번 학생의 근로시간")), false);
});
