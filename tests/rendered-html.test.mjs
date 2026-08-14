import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

test("builds a static GitHub Pages entry", async () => {
  const html = await readFile(new URL("../dist/index.html", import.meta.url), "utf8");

  assert.match(html, /<title>근로시간표 만들기<\/title>/i);
  assert.match(html, /<div id="root"><\/div>/i);
  assert.match(html, /\/sw_timetable\/assets\/[^"']+\.js/i);
  assert.match(html, /\/sw_timetable\/assets\/[^"']+\.css/i);
  assert.doesNotMatch(html, /localhost|dist\/server|wrangler/i);
});

test("keeps the work schedule features in the client app", async () => {
  const [page, packageJson, viteConfig] = await Promise.all([
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
    readFile(new URL("../package.json", import.meta.url), "utf8"),
    readFile(new URL("../vite.config.ts", import.meta.url), "utf8"),
  ]);

  assert.match(page, /학생 시간표 사진을 한꺼번에 끌어다 놓으세요/);
  assert.match(page, /시간표 자동 생성/);
  assert.match(page, /수정 완료 · 대기표 재생성/);
  assert.match(page, /엑셀로 저장/);
  assert.match(page, /blocked\[availablePerson\]\[dayIndex\]\[slot\]/);
  assert.match(page, /hours\[availablePerson\]/);
  assert.match(page, /minimumAttendanceDays,\s+operating,/);
  assert.match(page, /\[3, 4, 5, 6, 7, 8, 9, 10\]/);
  assert.match(page, /요일별 근로 운영시간 입력표/);
  assert.match(page, /07:00 — 22:00/);
  assert.match(page, /대기표 수정 완료/);
  assert.match(page, /사전 고정 배정/);
  assert.match(page, /사전 배정 제외/);
  assert.match(page, /한 시간에 여러 학생을 체크할 수 있습니다/);
  assert.match(page, /자동 배정/);
  assert.match(page, /fixedAssignments/);
  assert.match(page, /toggleExcludedAssignment/);
  assert.match(page, /https:\/\/github\.com\/ssungcoding\/sw_timetable/);
  assert.match(page, /candidate !== workAssignments\[dayIndex\]\[slot\]/);
  assert.match(packageJson, /"build":\s*"vite build"/);
  assert.match(viteConfig, /base:\s*"\/sw_timetable\/"/);
});
