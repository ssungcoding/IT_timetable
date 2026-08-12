import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);

  return worker.fetch(
    new Request("http://localhost/", { headers: { accept: "text/html" } }),
    { ASSETS: { fetch: async () => new Response("Not found", { status: 404 }) } },
    { waitUntil() {}, passThroughOnException() {} },
  );
}

test("server-renders the work schedule builder", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);

  const html = await response.text();
  assert.match(html, /<title>근로시간표 만들기<\/title>/i);
  assert.match(html, /수업시간만 칠하면/);
  assert.match(html, /시간표 자동 생성/);
  assert.match(html, /학생별 최소 출근 요일/);
  assert.match(html, /학생 시간표 사진을 한꺼번에 끌어다 놓으세요/);
  assert.match(html, /학생 시간표 사진 5장 또는 6장 일괄 인식/);
  assert.match(html, /대기시간표/);
  assert.doesNotMatch(html, /codex-preview|react-loading-skeleton/);
});

test("removes starter-only code and metadata", async () => {
  const [page, layout, packageJson] = await Promise.all([
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/layout.tsx", import.meta.url), "utf8"),
    readFile(new URL("../package.json", import.meta.url), "utf8"),
  ]);

  assert.doesNotMatch(page, /_sites-preview|SkeletonPreview|codex-preview/);
  assert.match(page, /이 시간표 선택하고 수정/);
  assert.match(page, /등록된 학생 중.*학생은 없습니다/);
  assert.match(page, /수업 또는 일정이 있어 배정할 수 없습니다/);
  assert.match(page, /nativeEvent\.isComposing/);
  assert.match(page, /onCompositionEnd/);
  assert.match(page, /수정 완료 · 대기표 재생성/);
  assert.match(page, /근로표 수정 완료 후 자동으로 다시 배정됩니다/);
  assert.match(page, /ySplit:\s*2/);
  assert.match(page, /s:\s*\{ r:\s*6, c:\s*1 \}, e:\s*\{ r:\s*6, c:\s*5 \}/);
  assert.match(page, /const row = 2 \+ slot \+ \(slot >= 4 \? 1 : 0\)/);
  assert.match(page, /buildScheduleCandidates\(blocked, 20, minimumAttendanceDays\)/);
  assert.match(layout, /title:\s*"근로시간표 만들기"/);
  assert.doesNotMatch(packageJson, /react-loading-skeleton/);
  assert.match(packageJson, /xlsx-js-style/);
});
