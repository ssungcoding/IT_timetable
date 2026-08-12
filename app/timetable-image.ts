type RecognitionResult = {
  blocked: boolean[][];
  blockedSlots: number;
  confidence: "high" | "review";
};

type CellColor = {
  day: number;
  slot: number;
  red: number;
  green: number;
  blue: number;
};

const median = (values: number[]) =>
  values.sort((a, b) => a - b)[Math.floor(values.length / 2)];

function regularGrid(
  scores: number[],
  width: number,
  pointCount: number,
) {
  let best = { score: -1, spacing: 0, offset: 0 };
  const spacingMin = Math.round(width * 0.113);
  const spacingMax = Math.round(width * 0.123);
  const offsetMin = Math.round(width * 0.04);
  const offsetMax = Math.round(width * 0.09);

  for (let spacing = spacingMin; spacing <= spacingMax; spacing += 1) {
    for (let offset = offsetMin; offset <= offsetMax; offset += 1) {
      let score = 0;
      let validPoints = 0;
      for (let point = 0; point < pointCount; point += 1) {
        const expected = offset + point * spacing;
        if (expected >= scores.length) break;
        let localPeak = 0;
        for (let delta = -3; delta <= 3; delta += 1) {
          localPeak = Math.max(localPeak, scores[expected + delta] ?? 0);
        }
        score += localPeak;
        validPoints += 1;
      }
      if (validPoints === pointCount && score > best.score) {
        best = { score, spacing, offset };
      }
    }
  }

  if (best.spacing === 0) {
    throw new Error("시간 격자를 찾지 못했습니다.");
  }
  return best;
}

export async function recognizeTimetableImage(file: File): Promise<RecognitionResult> {
  const bitmap = await createImageBitmap(file);
  const scale = Math.min(1, 1200 / bitmap.width);
  const width = Math.round(bitmap.width * scale);
  const height = Math.round(bitmap.height * scale);
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) throw new Error("이미지를 읽을 수 없습니다.");
  context.drawImage(bitmap, 0, 0, width, height);
  bitmap.close();

  const pixels = context.getImageData(0, 0, width, height).data;
  const rowScores = Array(height).fill(0) as number[];
  for (let y = 1; y < height - 1; y += 1) {
    let changed = 0;
    let samples = 0;
    for (let x = 0; x < width; x += 4) {
      const above = ((y - 1) * width + x) * 4;
      const below = ((y + 1) * width + x) * 4;
      const difference =
        Math.abs(pixels[above] - pixels[below]) +
        Math.abs(pixels[above + 1] - pixels[below + 1]) +
        Math.abs(pixels[above + 2] - pixels[below + 2]);
      if (difference > 18) changed += 1;
      samples += 1;
    }
    rowScores[y] = changed / samples;
  }

  const availableGridLines = Math.max(8, Math.min(9, Math.floor(height / (width * 0.113))));
  const grid = regularGrid(rowScores, width, availableGridLines);
  const lastSampleY = grid.offset + 7.75 * grid.spacing;
  if (lastSampleY >= height) {
    throw new Error("사진에 17시까지의 시간표가 모두 보여야 합니다.");
  }

  const colors: CellColor[] = [];
  for (let day = 0; day < 5; day += 1) {
    for (let slot = 0; slot < 12; slot += 1) {
      const hour = slot < 4
        ? 10 + Math.floor(slot / 2)
        : 13 + Math.floor((slot - 4) / 2);
      const hourOffset = hour - 9 + (slot % 2) * 0.5;
      const centerY = grid.offset + (hourOffset + 0.25) * grid.spacing;
      const centerX = width * (0.15 + day * 0.19);
      const reds: number[] = [];
      const greens: number[] = [];
      const blues: number[] = [];

      for (let sampleY = -3; sampleY <= 3; sampleY += 1) {
        for (let sampleX = -4; sampleX <= 4; sampleX += 1) {
          const x = Math.max(0, Math.min(
            width - 1,
            Math.round(centerX + sampleX * width * 0.012),
          ));
          const y = Math.max(0, Math.min(
            height - 1,
            Math.round(centerY + sampleY * grid.spacing * 0.045),
          ));
          const index = (y * width + x) * 4;
          reds.push(pixels[index]);
          greens.push(pixels[index + 1]);
          blues.push(pixels[index + 2]);
        }
      }

      colors.push({
        day,
        slot,
        red: median(reds),
        green: median(greens),
        blue: median(blues),
      });
    }
  }

  const buckets = new Map<string, CellColor[]>();
  colors.forEach((color) => {
    const key = [color.red, color.green, color.blue]
      .map((value) => Math.round(value / 16))
      .join(",");
    buckets.set(key, [...(buckets.get(key) ?? []), color]);
  });
  const backgroundCells = [...buckets.values()]
    .sort((a, b) => b.length - a.length)[0];
  if (!backgroundCells || backgroundCells.length < 8) {
    throw new Error("빈 시간과 수업 블록을 구분하지 못했습니다.");
  }
  const background = backgroundCells.reduce(
    (sum, color) => ({
      red: sum.red + color.red / backgroundCells.length,
      green: sum.green + color.green / backgroundCells.length,
      blue: sum.blue + color.blue / backgroundCells.length,
    }),
    { red: 0, green: 0, blue: 0 },
  );

  const blocked = Array.from({ length: 5 }, () => Array(12).fill(false) as boolean[]);
  let blockedSlots = 0;
  colors.forEach((color) => {
    const distance = Math.hypot(
      color.red - background.red,
      color.green - background.green,
      color.blue - background.blue,
    );
    if (distance > 16) {
      blocked[color.day][color.slot] = true;
      blockedSlots += 1;
    }
  });

  if (blockedSlots === 0) {
    throw new Error("색칠된 수업 블록을 찾지 못했습니다.");
  }

  return {
    blocked,
    blockedSlots,
    confidence: grid.score / availableGridLines > 0.2 ? "high" : "review",
  };
}
