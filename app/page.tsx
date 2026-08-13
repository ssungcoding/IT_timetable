"use client";

import { Fragment, useEffect, useRef, useState } from "react";
import {
  buildScheduleCandidates,
  createEmptyBlocked,
  createEmptyOperatingGrid,
  DAYS,
  recalculateScheduleResult,
  TIMES,
  updateStandbyAssignment,
  type BlockedGrid,
  type OperatingGrid,
  type ScheduleResult,
} from "./scheduler";
import { recognizeTimetableImage } from "./timetable-image";

const COLORS = [
  "#F4C95D", "#9BCB68", "#77B7D7", "#E88AA8", "#B09AD9",
  "#65C3B2", "#F0A7B5", "#A8B9E8", "#D0A56B", "#85C6A1",
];
const DARK_COLORS = [
  "#6C4C00", "#345614", "#174F6A", "#6D2341", "#4E3574",
  "#14574D", "#762D40", "#334B83", "#684313", "#24583A",
];

const makeNames = (count: number, previous: string[] = []) =>
  Array.from({ length: count }, (_, index) => previous[index] ?? `학생 ${index + 1}`);

function excelCell(sheet: Record<string, unknown>, address: string) {
  return sheet[address] as { s?: Record<string, unknown>; v?: unknown } | undefined;
}

async function downloadWorkbook(
  names: string[],
  result: ScheduleResult,
  candidateNumber: number,
  minimumAttendanceDays: number,
  operating: OperatingGrid,
) {
  const XLSX = await import("xlsx-js-style");
  const titleRows: (string | number)[][] = [
    ["근로 시간표", ...DAYS.map(() => "")],
    ["시간", ...DAYS],
  ];
  const scheduleRows: (string | number)[][] = [];
  const lastScheduleSheetRow = 1 + TIMES.length;

  TIMES.forEach((time, slot) => {
    scheduleRows.push([
      time,
      ...DAYS.map((_, day) => {
        if (!operating[day][slot]) return "운영 없음";
        const person = result.assignments[day][slot];
        return person === null ? "" : names[person];
      }),
    ]);
  });

  const scheduleSheet = XLSX.utils.aoa_to_sheet([...titleRows, ...scheduleRows]);
  const ws = scheduleSheet as unknown as Record<string, unknown> & {
    "!merges"?: { s: { r: number; c: number }; e: { r: number; c: number } }[];
    "!cols"?: { wch: number }[];
    "!rows"?: { hpt: number }[];
    "!freeze"?: { xSplit: number; ySplit: number };
  };
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: DAYS.length } },
  ];
  ws["!cols"] = [{ wch: 18 }, ...DAYS.map(() => ({ wch: 16 }))];
  ws["!rows"] = [{ hpt: 32 }, { hpt: 26 }];
  ws["!freeze"] = { xSplit: 1, ySplit: 2 };

  const border = {
    top: { style: "thin", color: { rgb: "AEB8AE" } },
    bottom: { style: "thin", color: { rgb: "AEB8AE" } },
    left: { style: "thin", color: { rgb: "AEB8AE" } },
    right: { style: "thin", color: { rgb: "AEB8AE" } },
  };

  for (let row = 0; row <= lastScheduleSheetRow; row += 1) {
    for (let col = 0; col <= DAYS.length; col += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = excelCell(ws, address);
      if (!cell) continue;
      cell.s = {
        font: { name: "맑은 고딕", sz: 11, color: { rgb: "18332B" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: row >= 1 ? border : undefined,
      };
    }
  }

  Object.assign(excelCell(ws, "A1")!.s!, {
    font: { name: "맑은 고딕", sz: 18, bold: true, color: { rgb: "FFFFFF" } },
    fill: { fgColor: { rgb: "173F35" } },
  });
  for (let col = 0; col <= DAYS.length; col += 1) {
    const cell = excelCell(ws, XLSX.utils.encode_cell({ r: 1, c: col }));
    if (cell) cell.s = {
      font: { name: "맑은 고딕", sz: 11, bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "2B6656" } },
      alignment: { horizontal: "center", vertical: "center" },
      border,
    };
  }

  for (let row = 2; row <= lastScheduleSheetRow; row += 1) {
    const timeCell = excelCell(ws, XLSX.utils.encode_cell({ r: row, c: 0 }));
    if (timeCell) timeCell.s = {
      font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "40534D" } },
      fill: { fgColor: { rgb: "F5F6F2" } },
      alignment: { horizontal: "center", vertical: "center" },
      border,
    };
    for (let col = 1; col <= DAYS.length; col += 1) {
      const cell = excelCell(ws, XLSX.utils.encode_cell({ r: row, c: col }));
      if (!cell) continue;
      const person = names.indexOf(String(cell.v));
      if (person === -1) {
        cell.s = {
          font: { name: "맑은 고딕", sz: 10, color: { rgb: "7A847E" } },
          fill: { fgColor: { rgb: "EEF0EC" } },
          alignment: { horizontal: "center", vertical: "center" },
          border,
        };
        continue;
      }
      cell.s = {
        font: { name: "맑은 고딕", sz: 11, bold: true, color: { rgb: DARK_COLORS[person].slice(1) } },
        fill: { fgColor: { rgb: COLORS[person].slice(1) } },
        alignment: { horizontal: "center", vertical: "center" },
        border,
      };
    }
  }

  const standbySheet = structuredClone(scheduleSheet);
  const standby = standbySheet as unknown as Record<string, unknown>;
  const standbyTitle = excelCell(standby, "A1");
  if (standbyTitle) standbyTitle.v = "대기 시간표";
  TIMES.forEach((_, slot) => {
    const row = 2 + slot;
    DAYS.forEach((__, day) => {
      const cell = excelCell(standby, XLSX.utils.encode_cell({ r: row, c: day + 1 }));
      if (!cell) return;
      if (!operating[day][slot]) {
        cell.v = "운영 없음";
        cell.s = {
          font: { name: "맑은 고딕", sz: 10, color: { rgb: "7A847E" } },
          fill: { fgColor: { rgb: "EEF0EC" } },
          alignment: { horizontal: "center", vertical: "center" },
          border,
        };
        return;
      }
      const person = result.standbyAssignments[day][slot];
      cell.v = person === null ? "" : names[person];
      cell.s = person === null ? {
        font: { name: "맑은 고딕", sz: 10, color: { rgb: "98A19B" } },
        fill: { fgColor: { rgb: "FFFFFF" } },
        alignment: { horizontal: "center", vertical: "center" },
        border,
      } : {
        font: { name: "맑은 고딕", sz: 11, bold: true, color: { rgb: DARK_COLORS[person].slice(1) } },
        fill: { fgColor: { rgb: COLORS[person].slice(1) } },
        alignment: { horizontal: "center", vertical: "center" },
        border,
      };
    });
  });

  const summaryRows: (string | number)[][] = [
    [`개인별 근로 요약 · 후보 ${candidateNumber}안 · 최소 ${minimumAttendanceDays}일 출근`, "", "", "", "", ""],
    ["이름", "주간 근로시간", "출근 요일 수", "근로 요일", "주간 대기시간", "기준 충족"],
    ...names.map((name, person) => [
      name,
      result.hours[person],
      result.attendanceDays[person],
      result.workDays[person].join(", "),
      result.standbyHours[person],
      result.hours[person] >= 4 && result.hours[person] <= 6 ? "충족" : "확인 필요",
    ]),
  ];
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
  const summary = summarySheet as unknown as Record<string, unknown> & {
    "!merges"?: { s: { r: number; c: number }; e: { r: number; c: number } }[];
    "!cols"?: { wch: number }[];
  };
  summary["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }];
  summary["!cols"] = [{ wch: 16 }, { wch: 17 }, { wch: 16 }, { wch: 22 }, { wch: 17 }, { wch: 14 }];

  for (let row = 0; row < summaryRows.length; row += 1) {
    for (let col = 0; col < 6; col += 1) {
      const cell = excelCell(summary, XLSX.utils.encode_cell({ r: row, c: col }));
      if (!cell) continue;
      cell.s = {
        font: { name: "맑은 고딕", sz: row === 0 ? 16 : 11, bold: row <= 1, color: { rgb: row === 0 ? "FFFFFF" : "18332B" } },
        fill: { fgColor: { rgb: row === 0 ? "173F35" : row === 1 ? "DCE8E1" : row % 2 ? "F7F8F5" : "FFFFFF" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: row > 0 ? border : undefined,
        numFmt: (col === 1 || col === 4) && row > 1 ? '0.0"시간"' : undefined,
      };
    }
  }

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, scheduleSheet, "근로시간표");
  XLSX.utils.book_append_sheet(workbook, standbySheet, "대기시간표");
  XLSX.utils.book_append_sheet(workbook, summarySheet, "개인별 요약");
  const bytes = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `근로시간표_후보${candidateNumber}_${new Date().toISOString().slice(0, 10)}.xlsx`;
  anchor.click();
  URL.revokeObjectURL(url);
}

function ScheduleTable({
  assignments,
  names,
  operating,
  emptyLabel = "—",
}: {
  assignments: (number | null)[][];
  names: string[];
  operating: OperatingGrid;
  emptyLabel?: string;
}) {
  return (
    <div className="result-table-wrap">
      <table className="result-table">
        <thead>
          <tr><th>시간</th>{DAYS.map((day) => <th key={day}>{day}요일</th>)}</tr>
        </thead>
        <tbody>
          {TIMES.map((time, slot) => (
            <Fragment key={time}>
              <tr>
                <th>{time}</th>
                {DAYS.map((day, dayIndex) => {
                  const person = assignments[dayIndex][slot];
                  if (!operating[dayIndex][slot]) {
                    return <td className="closed-assignment" key={day}>운영 안 함</td>;
                  }
                  return person === null ? (
                    <td className="empty-assignment" key={day}>{emptyLabel}</td>
                  ) : (
                    <td key={day} style={{ background: COLORS[person], color: DARK_COLORS[person] }}>
                      {names[person]}
                    </td>
                  );
                })}
              </tr>
            </Fragment>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function EditableScheduleTable({
  assignments,
  names,
  operating,
  blocked,
  hours,
  onSelect,
}: {
  assignments: (number | null)[][];
  names: string[];
  operating: OperatingGrid;
  blocked: BlockedGrid;
  hours: number[];
  onSelect: (day: number, slot: number, person: number) => void;
}) {
  return (
    <div className="result-table-wrap editing-table-wrap">
      <table className="result-table editing-table">
        <thead>
          <tr><th>시간</th>{DAYS.map((day) => <th key={day}>{day}요일</th>)}</tr>
        </thead>
        <tbody>
          {TIMES.map((time, slot) => (
            <Fragment key={time}>
              <tr>
                <th>{time}</th>
                {DAYS.map((day, dayIndex) => {
                  const person = assignments[dayIndex][slot];
                  if (person === null || !operating[dayIndex][slot]) {
                    return <td className="closed-assignment" key={day}>운영 안 함</td>;
                  }
                  return (
                    <td key={day} style={{ background: COLORS[person] }}>
                      <select
                        aria-label={`${day}요일 ${time} 근로 학생 수정`}
                        value={person}
                        onChange={(event) =>
                          onSelect(dayIndex, slot, Number(event.currentTarget.value))
                        }
                      >
                        {names.map((name, availablePerson) =>
                          blocked[availablePerson][dayIndex][slot] ? null : (
                            <option value={availablePerson} key={availablePerson}>
                              {name} · {hours[availablePerson]}시간
                            </option>
                          ),
                        )}
                      </select>
                    </td>
                  );
                })}
              </tr>
            </Fragment>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function EditableStandbyTable({
  assignments,
  workAssignments,
  names,
  operating,
  blocked,
  hours,
  onSelect,
}: {
  assignments: (number | null)[][];
  workAssignments: (number | null)[][];
  names: string[];
  operating: OperatingGrid;
  blocked: BlockedGrid;
  hours: number[];
  onSelect: (day: number, slot: number, person: number) => void;
}) {
  return (
    <div className="result-table-wrap editing-table-wrap standby-editing-table-wrap">
      <table className="result-table editing-table">
        <thead>
          <tr><th>시간</th>{DAYS.map((day) => <th key={day}>{day}요일</th>)}</tr>
        </thead>
        <tbody>
          {TIMES.map((time, slot) => (
            <Fragment key={time}>
              <tr>
                <th>{time}</th>
                {DAYS.map((day, dayIndex) => {
                  if (!operating[dayIndex][slot]) {
                    return <td className="closed-assignment" key={day}>운영 안 함</td>;
                  }
                  const person = assignments[dayIndex][slot];
                  const availablePeople = names.map((_, index) => index).filter(
                    (candidate) =>
                      candidate !== workAssignments[dayIndex][slot] &&
                      !blocked[candidate][dayIndex][slot],
                  );
                  if (availablePeople.length === 0) {
                    return <td className="empty-assignment" key={day}>가능한 학생 없음</td>;
                  }
                  const displayPerson = person !== null && availablePeople.includes(person)
                    ? person
                    : null;
                  return (
                    <td
                      key={day}
                      style={{ background: displayPerson === null ? undefined : COLORS[displayPerson] }}
                    >
                      <select
                        aria-label={`${day}요일 ${time} 대기 학생 수정`}
                        value={displayPerson ?? ""}
                        onChange={(event) =>
                          onSelect(dayIndex, slot, Number(event.currentTarget.value))
                        }
                      >
                        {displayPerson === null && <option value="" disabled>빈칸 · 학생 선택</option>}
                        {availablePeople.map((availablePerson) => (
                          <option value={availablePerson} key={availablePerson}>
                            {names[availablePerson]} · {hours[availablePerson]}시간
                          </option>
                        ))}
                      </select>
                    </td>
                  );
                })}
              </tr>
            </Fragment>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export default function Home() {
  const [peopleCount, setPeopleCount] = useState(5);
  const [minimumAttendanceDays, setMinimumAttendanceDays] = useState(1);
  const [names, setNames] = useState(() => makeNames(5));
  const [blocked, setBlocked] = useState<BlockedGrid>(() => createEmptyBlocked(5));
  const [operating, setOperating] = useState<OperatingGrid>(() => createEmptyOperatingGrid());
  const [operatingPaint, setOperatingPaint] = useState<boolean | null>(null);
  const [paintValue, setPaintValue] = useState<{ person: number; value: boolean } | null>(null);
  const [results, setResults] = useState<ScheduleResult[]>([]);
  const [selectedCandidate, setSelectedCandidate] = useState(0);
  const [editedResult, setEditedResult] = useState<ScheduleResult | null>(null);
  const [isFinalEditing, setIsFinalEditing] = useState(false);
  const [isStandbyEditing, setIsStandbyEditing] = useState(false);
  const [recognitionStatus, setRecognitionStatus] = useState(() => Array(5).fill(""));
  const [batchDragging, setBatchDragging] = useState(false);
  const [batchStatus, setBatchStatus] = useState("");
  const [error, setError] = useState("");
  const resultRef = useRef<HTMLElement>(null);

  useEffect(() => {
    const stopPainting = () => {
      setPaintValue(null);
      setOperatingPaint(null);
    };
    window.addEventListener("pointerup", stopPainting);
    window.addEventListener("pointercancel", stopPainting);
    return () => {
      window.removeEventListener("pointerup", stopPainting);
      window.removeEventListener("pointercancel", stopPainting);
    };
  }, []);

  const clearGeneratedResults = () => {
    setResults([]);
    setEditedResult(null);
    setIsFinalEditing(false);
    setIsStandbyEditing(false);
  };

  const changeCount = (count: number) => {
    if (count === peopleCount) return;
    setPeopleCount(count);
    setNames((current) => makeNames(count, current));
    setBlocked((current) => {
      const next = createEmptyBlocked(count);
      for (let person = 0; person < Math.min(current.length, count); person += 1) {
        next[person] = current[person].map((day) => [...day]);
      }
      return next;
    });
    setRecognitionStatus((current) =>
      Array.from({ length: count }, (_, index) => current[index] ?? ""),
    );
    clearGeneratedResults();
    setError("");
  };

  const setCell = (targetPerson: number, day: number, slot: number, value: boolean) => {
    setBlocked((current) =>
      current.map((personGrid, person) =>
        person !== targetPerson
          ? personGrid
          : personGrid.map((dayGrid, dayIndex) =>
              dayIndex !== day
                ? dayGrid
                : dayGrid.map((cell, slotIndex) =>
                    slotIndex === slot ? value : cell,
                  ),
            ),
      ),
    );
    clearGeneratedResults();
  };

  const setOperatingCell = (day: number, slot: number, value: boolean) => {
    setOperating((current) => current.map((dayGrid, dayIndex) =>
      dayIndex === day
        ? dayGrid.map((cell, slotIndex) => slotIndex === slot ? value : cell)
        : dayGrid,
    ));
    clearGeneratedResults();
    setError("");
  };

  const generate = () => {
    const cleanNames = names.map((name, index) => name.trim() || `학생 ${index + 1}`);
    setNames(cleanNames);
    setError("");
    setEditedResult(null);
    setIsFinalEditing(false);
    setIsStandbyEditing(false);
    try {
      const nextResults = buildScheduleCandidates(
        blocked,
        20,
        minimumAttendanceDays,
        operating,
      );
      setResults(nextResults);
      setSelectedCandidate(0);
      requestAnimationFrame(() =>
        resultRef.current?.scrollIntoView({ behavior: "smooth", block: "start" }),
      );
    } catch (reason) {
      clearGeneratedResults();
      setError(reason instanceof Error ? reason.message : "시간표를 만들지 못했습니다.");
    }
  };

  const recognizeImage = async (person: number, file: File) => {
    if (!file.type.startsWith("image/")) {
      setRecognitionStatus((current) => current.map((status, index) =>
        index === person ? "인식 실패 · 이미지 파일만 넣어주세요." : status,
      ));
      return false;
    }
    setRecognitionStatus((current) => current.map((status, index) =>
      index === person ? "사진을 분석하고 있어요…" : status,
    ));
    try {
      const recognized = await recognizeTimetableImage(file);
      const fileName = file.name.replace(/\.[^/.]+$/, "").normalize("NFC").trim();
      if (fileName) {
        setNames((current) => current.map((name, index) =>
          index === person ? fileName : name,
        ));
      }
      setBlocked((current) => current.map((personGrid, index) =>
        index === person ? recognized.blocked : personGrid,
      ));
      clearGeneratedResults();
      setError("");
      setRecognitionStatus((current) => current.map((status, index) =>
        index === person
          ? `${fileName || "사진"} · ${recognized.blockedSlots}칸 인식 완료 · 표에서 바로 수정할 수 있어요.`
          : status,
      ));
      return true;
    } catch (reason) {
      const message = reason instanceof Error ? reason.message : "사진을 인식하지 못했습니다.";
      setRecognitionStatus((current) => current.map((status, index) =>
        index === person ? `인식 실패 · ${message}` : status,
      ));
      return false;
    }
  };

  const recognizeImageBatch = async (files: File[]) => {
    const imageFiles = files.filter((file) => file.type.startsWith("image/"));
    if (imageFiles.length < 3 || imageFiles.length > 10) {
      setBatchStatus("시간표 사진을 3장부터 10장까지 한꺼번에 넣어주세요.");
      return;
    }

    changeCount(imageFiles.length);
    setBatchStatus(`${imageFiles.length}장의 시간표를 동시에 분석하고 있어요…`);
    const recognitionResults = await Promise.all(
      imageFiles.map((file, person) => recognizeImage(person, file)),
    );
    const successCount = recognitionResults.filter(Boolean).length;
    setBatchStatus(
      successCount === imageFiles.length
        ? `${successCount}명 일괄 인식 완료 · 각 표를 확인해 주세요.`
        : `${successCount}/${imageFiles.length}명 인식 완료 · 실패한 학생의 사진을 확인해 주세요.`,
    );
  };

  const beginFinalEditing = (candidate: ScheduleResult) => {
    setEditedResult(structuredClone(candidate));
    setIsFinalEditing(true);
    setIsStandbyEditing(false);
  };

  const selectManualAssignment = (day: number, slot: number, person: number) => {
    if (!editedResult || !operating[day][slot]) return;
    if (blocked[person][day][slot]) {
      window.alert(`${names[person]} 학생은 ${DAYS[day]}요일 ${TIMES[slot]}에 수업 또는 일정이 있어 배정할 수 없습니다.`);
      return;
    }

    const assignments = editedResult.assignments.map((dayAssignments) => [...dayAssignments]);
    assignments[day][slot] = person;
    const nextResult = recalculateScheduleResult(
      assignments,
      blocked,
      minimumAttendanceDays,
      operating,
    );
    setEditedResult(nextResult);
  };

  const finishFinalEditing = () => {
    if (!editedResult) return;
    setEditedResult(recalculateScheduleResult(
      editedResult.assignments,
      blocked,
      minimumAttendanceDays,
      operating,
    ));
    setIsFinalEditing(false);
    setIsStandbyEditing(false);
  };

  const selectManualStandby = (day: number, slot: number, person: number) => {
    if (!editedResult || !operating[day][slot]) return;
    if (
      person === editedResult.assignments[day][slot] ||
      blocked[person][day][slot]
    ) {
      window.alert(`${names[person]} 학생은 이 시간의 대기 근로자로 배정할 수 없습니다.`);
      return;
    }

    setEditedResult(updateStandbyAssignment(
      editedResult,
      blocked,
      operating,
      day,
      slot,
      person,
    ));
  };

  const candidateResult = results[selectedCandidate] ?? null;
  const result = editedResult ?? candidateResult;

  return (
    <main>
      <header className="topbar">
        <a className="brand" href="#top" aria-label="근로시간표 만들기 처음으로">
          <span className="brand-mark" aria-hidden="true">W</span>
          <span>근로시간표 만들기</span>
        </a>
        <span className="topbar-note">월–일 · 30분 단위</span>
      </header>

      <section className="hero" id="top">
        <div>
          <p className="eyebrow">WORK SCHEDULE BUILDER</p>
          <h1>SW교육원 근로시간표 편성</h1>
        </div>
        <div className="operation-card" aria-label="운영 기준">
          <div className="operation-label">운영 시간</div>
          <strong>07:00 — 22:00</strong>
          <div className="operation-divider" />
          <span>30분 단위로 직접 운영시간 선택</span>
        </div>
      </section>

      <section className="workspace" aria-label="근로시간표 입력">
        <div className="section-heading">
          <span className="step-number">01</span>
          <div>
            <h2>근로 운영시간을 색칠하세요</h2>
            <p>색칠한 요일과 시간에만 근로 및 대기 인원을 편성합니다.</p>
          </div>
        </div>

        <article className="availability-card operating-card">
          <div className="grid-toolbar">
            <div className="student-card-title">
              <span className="operating-mark" aria-hidden="true">W</span>
              <div>
                <strong>요일별 운영시간</strong>
                <small>{operating.flat().filter(Boolean).length * 0.5}시간 선택</small>
              </div>
            </div>
            <div className="card-actions">
              <button
                type="button"
                className="clear-button"
                onClick={() => {
                  setOperating(DAYS.map(() => TIMES.map(() => true)));
                  clearGeneratedResults();
                  setError("");
                }}
              >
                전체 선택
              </button>
              <button
                type="button"
                className="clear-button"
                onClick={() => {
                  setOperating(DAYS.map(() => TIMES.map(() => false)));
                  clearGeneratedResults();
                  setError("");
                }}
              >
                전체 지우기
              </button>
            </div>
          </div>
          <div className="table-scroll">
            <div className="availability-grid operating-grid" role="grid" aria-label="요일별 근로 운영시간 입력표">
              <div className="grid-corner">시간</div>
              {DAYS.map((day) => <div className="day-header" key={day}>{day}요일</div>)}
              {TIMES.map((time, slot) => (
                <div className="grid-row" key={time}>
                  <div className="time-label">{time}</div>
                  {DAYS.map((day, dayIndex) => {
                    const isOperating = operating[dayIndex][slot];
                    return (
                      <button
                        key={day}
                        type="button"
                        aria-label={`${day}요일 ${time} ${isOperating ? "근로 운영" : "운영 안 함"}`}
                        aria-pressed={isOperating}
                        className={`availability-cell operating-cell ${isOperating ? "selected" : ""}`}
                        onPointerDown={(event) => {
                          event.preventDefault();
                          const value = !isOperating;
                          setOperatingPaint(value);
                          setOperatingCell(dayIndex, slot, value);
                        }}
                        onPointerEnter={() => {
                          if (operatingPaint !== null && operatingPaint !== isOperating) {
                            setOperatingCell(dayIndex, slot, operatingPaint);
                          }
                        }}
                        onKeyDown={(event) => {
                          if (event.key === "Enter" || event.key === " ") {
                            event.preventDefault();
                            setOperatingCell(dayIndex, slot, !isOperating);
                          }
                        }}
                      >
                        {isOperating && <span>운영</span>}
                      </button>
                    );
                  })}
                </div>
              ))}
            </div>
          </div>
          <div className="legend operating-legend">
            <span><i className="legend-operating" /> 근로 편성</span>
            <span><i className="legend-closed" /> 운영 안 함</span>
            <span className="legend-tip">필요한 요일과 시간만 직접 색칠하세요.</span>
          </div>
        </article>

        <div className="section-heading">
          <span className="step-number">02</span>
          <div>
            <h2>학생 정보를 입력하세요</h2>
            <p>근로 인원을 고르고 이름을 확인해 주세요.</p>
          </div>
        </div>

        <div className="setup-panel">
          <div className="setup-options">
            <div className="count-control">
              <span className="field-label">근로 인원</span>
              <div className="segment-control count-range" role="group" aria-label="근로 인원 선택">
                {[3, 4, 5, 6, 7, 8, 9, 10].map((count) => (
                  <button
                    key={count}
                    className={peopleCount === count ? "active" : ""}
                    onClick={() => changeCount(count)}
                    type="button"
                  >
                    {count}명
                  </button>
                ))}
              </div>
            </div>
            <div className="count-control">
              <span className="field-label">학생별 최소 출근 요일</span>
              <div className="segment-control day-control" role="group" aria-label="학생별 최소 출근 요일 선택">
                {[1, 2, 3, 4, 5, 6, 7].map((days) => (
                  <button
                    key={days}
                    className={minimumAttendanceDays === days ? "active" : ""}
                    onClick={() => {
                      setMinimumAttendanceDays(days);
                      clearGeneratedResults();
                      setError("");
                    }}
                    type="button"
                    aria-label={`최소 ${days}일 출근`}
                  >
                    {days}일
                  </button>
                ))}
              </div>
              <small className="field-help">각 학생이 월~금 중 최소 {minimumAttendanceDays}일 출근하며, 필요하면 더 많은 요일로 자동 배정됩니다.</small>
            </div>
          </div>
          <div className="name-grid">
            {names.map((name, index) => (
              <label className="name-field" key={index}>
                <span style={{ background: COLORS[index] }}>{index + 1}</span>
                <input
                  aria-label={`${index + 1}번 학생 이름`}
                  value={name}
                  onChange={(event) => {
                    const value = event.target.value;
                    setNames((current) => current.map((item, person) => person === index ? value : item));
                    clearGeneratedResults();
                  }}
                />
              </label>
            ))}
          </div>
        </div>

        <div className="section-heading schedule-heading">
          <span className="step-number">03</span>
          <div>
            <h2>수업시간을 색칠하세요</h2>
            <p>사진 3~10장을 한꺼번에 넣어 자동 인식하거나, 필요한 칸을 직접 수정하세요.</p>
          </div>
        </div>

        <label
          className={`batch-drop-zone ${batchDragging ? "dragging" : ""}`}
          onDragEnter={(event) => {
            if (event.dataTransfer.types.includes("Files")) {
              event.preventDefault();
              setBatchDragging(true);
            }
          }}
          onDragOver={(event) => {
            if (event.dataTransfer.types.includes("Files")) {
              event.preventDefault();
              event.dataTransfer.dropEffect = "copy";
              setBatchDragging(true);
            }
          }}
          onDragLeave={(event) => {
            const nextTarget = event.relatedTarget;
            if (!(nextTarget instanceof Node) || !event.currentTarget.contains(nextTarget)) {
              setBatchDragging(false);
            }
          }}
          onDrop={async (event) => {
            event.preventDefault();
            setBatchDragging(false);
            await recognizeImageBatch(Array.from(event.dataTransfer.files));
          }}
        >
          <input
            type="file"
            accept="image/jpeg,image/png,image/webp"
            multiple
            aria-label="학생 시간표 사진 3장부터 10장 일괄 인식"
            onChange={async (event) => {
              await recognizeImageBatch(Array.from(event.target.files ?? []));
              event.target.value = "";
            }}
          />
          <span className="batch-drop-icon" aria-hidden="true">⇩</span>
          <span className="batch-drop-copy">
            <strong>{batchDragging ? "사진을 여기에 놓으세요" : "학생 시간표 사진을 한꺼번에 끌어다 놓으세요"}</strong>
            <small>3~10장 · 사진 수에 맞춰 인원과 파일명을 자동 입력합니다.</small>
          </span>
          <span className="batch-select-button">여러 장 선택</span>
        </label>
        {batchStatus && <div className="batch-status" role="status">{batchStatus}</div>}

        <div className="availability-cards">
          {names.map((name, person) => {
            const classHours = blocked[person].flat().filter(Boolean).length * 0.5;
            return (
              <article className="availability-card" key={person}>
                <div className="grid-toolbar">
                  <div className="student-card-title">
                    <span style={{ background: COLORS[person] }}>{person + 1}</span>
                    <div>
                      <strong>{name.trim() || `학생 ${person + 1}`}</strong>
                      <small>수업 / 일정 {classHours}시간 선택</small>
                    </div>
                  </div>
                  <div className="card-actions">
                    <button
                      type="button"
                      className="clear-button"
                      onClick={() => {
                        setBlocked((current) => current.map((personGrid, index) =>
                          index === person ? personGrid.map((day) => day.map(() => false)) : personGrid,
                        ));
                        setRecognitionStatus((current) => current.map((status, index) =>
                          index === person ? "" : status,
                        ));
                        clearGeneratedResults();
                      }}
                    >
                      전체 지우기
                    </button>
                  </div>
                </div>

                {recognitionStatus[person] && (
                  <div className={`recognition-status ${recognitionStatus[person].startsWith("인식 실패") ? "failed" : ""}`}>
                    {recognitionStatus[person]}
                  </div>
                )}

                <div className="table-scroll">
                  <div className="availability-grid" role="grid" aria-label={`${name} 수업시간 입력표`}>
                    <div className="grid-corner">시간</div>
                    {DAYS.map((day) => <div className="day-header" key={day}>{day}요일</div>)}
                    {TIMES.map((time, slot) => (
                      <div className="grid-row" key={time}>
                        <div className="time-label">{time}</div>
                        {DAYS.map((day, dayIndex) => {
                          const isBlocked = blocked[person][dayIndex][slot];
                          return (
                            <button
                              key={day}
                              type="button"
                              aria-label={`${name} ${day}요일 ${time} ${isBlocked ? "수업 있음" : "근로 가능"}`}
                              aria-pressed={isBlocked}
                              className={`availability-cell ${isBlocked ? "blocked" : ""}`}
                              onPointerDown={(event) => {
                                event.preventDefault();
                                const value = !isBlocked;
                                setPaintValue({ person, value });
                                setCell(person, dayIndex, slot, value);
                              }}
                              onPointerEnter={() => {
                                if (paintValue?.person === person && paintValue.value !== isBlocked) {
                                  setCell(person, dayIndex, slot, paintValue.value);
                                }
                              }}
                              onKeyDown={(event) => {
                                if (event.key === "Enter" || event.key === " ") {
                                  event.preventDefault();
                                  setCell(person, dayIndex, slot, !isBlocked);
                                }
                              }}
                            >
                              {isBlocked && <span>수업</span>}
                            </button>
                          );
                        })}
                      </div>
                    ))}
                  </div>
                </div>
              </article>
            );
          })}
          <div className="legend global-legend">
            <span><i className="legend-available" /> 근로 가능</span>
            <span><i className="legend-blocked" /> 수업 / 일정</span>
            <span className="legend-tip">팁: 누른 채로 움직이면 여러 칸을 한 번에 칠할 수 있어요.</span>
          </div>
        </div>

        <div className="generate-row">
          <div>
            <strong>입력을 마쳤나요?</strong>
            <span>비어 있는 칸은 근로 가능한 시간으로 계산됩니다.</span>
          </div>
          <button className="generate-button" type="button" onClick={generate}>
            시간표 자동 생성 <span aria-hidden="true">→</span>
          </button>
        </div>
        {error && <div className="error-message" role="alert">{error}</div>}
      </section>

      <section className={`results ${result ? "visible" : ""}`} ref={resultRef}>
        <div className="results-inner">
          <div className="results-heading">
            <div className="section-heading">
              <span className="step-number light">04</span>
              <div>
                <h2>추천 시간표 {results.length}개</h2>
                <p>근로시간 형평성, 출근 요일 형평성 순으로 정렬했습니다.</p>
              </div>
            </div>
            {result && (
              <button
                className="download-button"
                type="button"
                disabled={isFinalEditing || isStandbyEditing}
                title={isFinalEditing || isStandbyEditing ? "수정 완료 후 저장할 수 있습니다." : undefined}
                onClick={() => downloadWorkbook(
                  names,
                  result,
                  selectedCandidate + 1,
                  minimumAttendanceDays,
                  operating,
                )}
              >
                <span aria-hidden="true">↓</span> 엑셀로 저장
              </button>
            )}
          </div>

          {!result ? (
            <div className="empty-result">
              <span aria-hidden="true">▦</span>
              <strong>아직 생성된 시간표가 없습니다.</strong>
              <p>학생별 수업시간을 입력한 뒤 자동 생성 버튼을 눌러주세요.</p>
            </div>
          ) : (
            <>
              <div className="candidate-picker" aria-label="시간표 후보 선택">
                {results.map((_, index) => (
                  <button
                    type="button"
                    key={index}
                    className={selectedCandidate === index ? "active" : ""}
                    onClick={() => {
                      setSelectedCandidate(index);
                      setEditedResult(null);
                      setIsFinalEditing(false);
                      setIsStandbyEditing(false);
                    }}
                  >
                    <span>{index + 1}</span>안
                  </button>
                ))}
              </div>
              <div className="ranking-note">
                현재 <strong>{selectedCandidate + 1}안{editedResult ? " · 최종 수정본" : ""}</strong> · 학생별 최소 {minimumAttendanceDays}일 출근 · 주간 근로시간 차이 → 출근 요일 수 차이 → 전체 출근 횟수 순
              </div>
              <div className={`final-edit-row ${isFinalEditing ? "active" : editedResult ? "completed" : ""}`}>
                <div>
                  <strong>{isFinalEditing ? "선택한 시간표를 수정하고 있습니다." : editedResult ? "최종 수정이 완료되었습니다." : "이 후보를 최종 시간표로 사용할까요?"}</strong>
                  <span>{isFinalEditing ? "근로표를 확인한 뒤 수정 완료를 누르면 대기시간표가 자동으로 다시 배정됩니다." : editedResult ? "수정된 근로표, 대기표와 개인별 요약을 확인하고 엑셀로 저장할 수 있습니다." : "선택 후 근로표의 학생 이름을 직접 바꿀 수 있습니다."}</span>
                </div>
                {isFinalEditing ? (
                  <div className="final-edit-actions">
                    <button
                      type="button"
                      className="reset-edit-button"
                      onClick={() => {
                        if (!candidateResult) return;
                        beginFinalEditing(candidateResult);
                      }}
                    >
                      후보 원본으로 되돌리기
                    </button>
                    <button type="button" className="finish-edit-button" onClick={finishFinalEditing}>
                      수정 완료 · 대기표 재생성
                    </button>
                  </div>
                ) : editedResult ? (
                  <button
                    type="button"
                    className="select-edit-button"
                    onClick={() => beginFinalEditing(editedResult)}
                  >
                    다시 수정
                  </button>
                ) : (
                  <button
                    type="button"
                    className="select-edit-button"
                    onClick={() => candidateResult && beginFinalEditing(candidateResult)}
                  >
                    이 시간표 선택하고 수정
                  </button>
                )}
              </div>
              {result.warnings.length > 0 && (
                <div className="warning-message">{result.warnings.join(" ")} 수업시간 제약으로 기준을 완전히 맞추지 못했습니다.</div>
              )}
              <div className="schedule-block">
                <div className="schedule-block-heading">
                  <div><span>MAIN</span><h3>근로시간표</h3></div>
                  <p>{isFinalEditing ? "역삼각형을 눌러 가능한 학생과 현재 근로시간을 확인하세요." : "각 시간대에 근로 학생 1명을 배정합니다."}</p>
                </div>
                {isFinalEditing && editedResult ? (
                  <EditableScheduleTable
                    assignments={editedResult.assignments}
                    names={names}
                    operating={operating}
                    blocked={blocked}
                    hours={editedResult.hours}
                    onSelect={selectManualAssignment}
                  />
                ) : (
                  <ScheduleTable assignments={result.assignments} names={names} operating={operating} />
                )}
              </div>

              <div className="schedule-block standby-block">
                <div className="schedule-block-heading">
                  <div>
                    <span>BACKUP</span><h3>대기시간표</h3>
                    {editedResult && !isFinalEditing && (
                      <button
                        type="button"
                        className={isStandbyEditing ? "finish-standby-button" : "edit-standby-button"}
                        onClick={() => setIsStandbyEditing((current) => !current)}
                      >
                        {isStandbyEditing ? "대기표 수정 완료" : "대기표 수정"}
                      </button>
                    )}
                  </div>
                  <p>{isFinalEditing
                    ? "근로표 수정 완료 후 자동으로 다시 배정됩니다."
                    : isStandbyEditing
                      ? "역삼각형을 눌러 가능한 학생과 현재 대기시간을 확인하세요."
                      : <>본 근로자를 제외한 가능한 학생을 공평하게 배정합니다.{result.unfilledStandby > 0 && ` 빈칸 ${result.unfilledStandby}개`}</>
                  }</p>
                </div>
                {isFinalEditing ? (
                  <div className="standby-pending-message">
                    <span aria-hidden="true">↻</span>
                    <strong>근로시간표를 먼저 완성해 주세요.</strong>
                    <p>`수정 완료 · 대기표 재생성`을 누르면 변경된 근로자를 제외하고 대기시간을 공정하게 다시 계산합니다.</p>
                  </div>
                ) : isStandbyEditing && editedResult ? (
                  <EditableStandbyTable
                    assignments={editedResult.standbyAssignments}
                    workAssignments={editedResult.assignments}
                    names={names}
                    operating={operating}
                    blocked={blocked}
                    hours={editedResult.standbyHours}
                    onSelect={selectManualStandby}
                  />
                ) : (
                  <ScheduleTable
                    assignments={result.standbyAssignments}
                    names={names}
                    operating={operating}
                    emptyLabel="빈칸"
                  />
                )}
              </div>

              <div className="summary-grid">
                {names.map((name, person) => (
                  <article className="summary-card" key={person}>
                    <div className="summary-person">
                      <span style={{ background: COLORS[person] }}>{name.slice(0, 1) || person + 1}</span>
                      <div><strong>{name}</strong><small>{result.workDays[person].map((day) => `${day}요일`).join(" · ")}</small></div>
                    </div>
                    <div className="summary-metrics">
                      <div><span>주간 근로</span><strong>{result.hours[person]}시간</strong></div>
                      <div><span>출근 요일</span><strong>{result.attendanceDays[person]}일</strong></div>
                      <div><span>주간 대기</span><strong>{result.standbyHours[person]}시간</strong></div>
                    </div>
                  </article>
                ))}
              </div>
            </>
          )}
        </div>
      </section>

      <footer>
        <div className="footer-brand">
          <strong>SW교육원 근로시간표 편성</strong>
          <span>VERSION 1.0.1</span>
        </div>
        <div className="footer-info">
          <span>Designed &amp; built by <strong>Sunghyun Kim</strong></span>
          <a href="mailto:atwoddl@naver.com">atwoddl@naver.com</a>
          <span>Created · 2026.08.12</span>
          <span>© 2026 Sunghyun Kim. All rights reserved.</span>
        </div>
      </footer>
    </main>
  );
}
