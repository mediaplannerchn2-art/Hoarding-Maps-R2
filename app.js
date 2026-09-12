(() => {
  'use strict';

  const NAVY = '#063B75';
  const RED = '#E00000';
  const BLACK = '#000000';
  const WHITE = '#FFFFFF';
  const GRID = '#CCD2D8';
  const TEXT = '#111111';
  const TABLE_COLUMNS = ['S.No', 'City', 'Media', 'Area Name', 'Location', 'W', 'H'];
  const CANONICAL_COLUMNS = [...TABLE_COLUMNS, 'Lat', 'Long'];

  const ALIASES = {
    'S.No': ['sno', 's no', 's.no', 'slno', 'sl no', 'serial no', 'serial number', 'no'],
    'City': ['city'],
    'Media': ['media', 'media type'],
    'Area Name': ['area name', 'area', 'location area'],
    'Location': ['location', 'site location', 'description'],
    'W': ['w', 'width', 'size w', 'size width', 'sizew'],
    'H': ['h', 'height', 'size h', 'size height', 'sizeh'],
    'Lat': ['lat', 'latitude'],
    'Long': ['long', 'lng', 'lon', 'longitude'],
  };

  const state = {
    workbook: null,
    rows: [],
    fileName: '',
    generated: false,
    lastBlob: null,
    tileCache: new Map(),
  };

  const $ = (id) => document.getElementById(id);
  const fileInput = $('fileInput');
  const dropZone = $('dropZone');
  const fileMeta = $('fileMeta');
  const sheetSection = $('sheetSection');
  const sheetSelect = $('sheetSelect');
  const settingsSection = $('settingsSection');
  const sizeSelect = $('sizeSelect');
  const highlightLast = $('highlightLast');
  const extraHighlights = $('extraHighlights');
  const generateBtn = $('generateBtn');
  const downloadBtn = $('downloadBtn');
  const statusBox = $('statusBox');
  const rowCount = $('rowCount');
  const renderInfo = $('renderInfo');
  const canvasStage = $('canvasStage');
  const emptyState = $('emptyState');
  const canvas = $('mapCanvas');
  const ctx = canvas.getContext('2d', { alpha: false });
  const dataPreview = $('dataPreview');
  const previewTable = $('previewTable');

  function showStatus(message, type = 'info') {
    statusBox.textContent = message;
    statusBox.className = `status-box ${type}`;
  }

  function hideStatus() {
    statusBox.className = 'status-box hidden';
    statusBox.textContent = '';
  }

  function normalizeName(value) {
    return String(value ?? '')
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  const aliasLookup = new Map();
  Object.entries(ALIASES).forEach(([canonical, names]) => {
    [...names, canonical].forEach((name) => aliasLookup.set(normalizeName(name), canonical));
  });

  function formatCell(value) {
    if (value === null || value === undefined) return '';
    if (typeof value === 'number' && Number.isInteger(value)) return String(value);
    return String(value).trim();
  }

  function parseNumber(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : NaN;
    const s = String(value ?? '').trim();
    if (!s) return NaN;
    return Number(s.replace(/,/g, ''));
  }

  function normalizeSheet(sheet) {
    const matrix = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: '',
      raw: true,
      blankrows: false,
    });

    if (!matrix.length) throw new Error('The selected sheet is empty.');

    const headerIndex = matrix.findIndex((row) => row.some((cell) => String(cell ?? '').trim() !== ''));
    if (headerIndex < 0) throw new Error('No header row was found in the selected sheet.');

    const headers = matrix[headerIndex].map((v) => String(v ?? '').trim());
    const selected = new Map();
    const canonicalByKey = new Map(CANONICAL_COLUMNS.map((c) => [normalizeName(c), c]));

    // Exact canonical headers first.
    headers.forEach((header, index) => {
      const canonical = canonicalByKey.get(normalizeName(header));
      if (canonical && !selected.has(canonical)) selected.set(canonical, index);
    });

    // Fill missing fields from aliases.
    headers.forEach((header, index) => {
      const canonical = aliasLookup.get(normalizeName(header));
      if (canonical && !selected.has(canonical)) selected.set(canonical, index);
    });

    const missing = CANONICAL_COLUMNS.filter((c) => !selected.has(c));
    if (missing.length) throw new Error(`Missing required columns: ${missing.join(', ')}`);

    const rows = [];
    for (let r = headerIndex + 1; r < matrix.length; r += 1) {
      const source = matrix[r];
      if (!source || source.every((cell) => String(cell ?? '').trim() === '')) continue;

      const row = {};
      CANONICAL_COLUMNS.forEach((col) => {
        row[col] = source[selected.get(col)] ?? '';
      });

      const rawLat = row.Lat;
      const rawLon = row.Long;
      row.Lat = parseNumber(row.Lat);
      row.Long = parseNumber(row.Long);

      if (!Number.isFinite(row.Lat) || !Number.isFinite(row.Long)) {
        throw new Error(
          `Invalid or missing GPS coordinate at S.No ${formatCell(row['S.No']) || r + 1}: ` +
          `Lat=${JSON.stringify(rawLat)}, Long=${JSON.stringify(rawLon)}`
        );
      }
      if (row.Lat < -90 || row.Lat > 90 || row.Long < -180 || row.Long > 180) {
        throw new Error(
          `GPS coordinate outside valid range at S.No ${formatCell(row['S.No']) || r + 1}: ` +
          `Lat=${row.Lat}, Long=${row.Long}`
        );
      }

      TABLE_COLUMNS.forEach((col) => {
        row[col] = row[col] ?? '';
      });
      rows.push(row);
    }

    if (!rows.length) throw new Error('No site rows were found in the selected sheet.');
    return rows;
  }

  function duplicateCoordinateGroups(rows) {
    const groups = new Map();
    rows.forEach((row) => {
      const key = `${row.Lat.toFixed(7)},${row.Long.toFixed(7)}`;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(formatCell(row['S.No']));
    });
    return [...groups.entries()]
      .filter(([, labels]) => labels.length > 1)
      .map(([key, labels]) => ({ key, labels }));
  }

  function updateDataPreview(rows) {
    const thead = previewTable.querySelector('thead');
    const tbody = previewTable.querySelector('tbody');
    thead.innerHTML = '';
    tbody.innerHTML = '';

    const trh = document.createElement('tr');
    CANONICAL_COLUMNS.forEach((col) => {
      const th = document.createElement('th');
      th.textContent = col;
      trh.appendChild(th);
    });
    thead.appendChild(trh);

    rows.slice(0, 100).forEach((row) => {
      const tr = document.createElement('tr');
      CANONICAL_COLUMNS.forEach((col) => {
        const td = document.createElement('td');
        td.textContent = formatCell(row[col]);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    dataPreview.classList.remove('hidden');
    rowCount.textContent = `${rows.length} site${rows.length === 1 ? '' : 's'}`;
  }

  async function handleFile(file) {
    if (!file) return;
    hideStatus();
    downloadBtn.classList.add('hidden');
    state.generated = false;
    state.lastBlob = null;

    try {
      if (typeof XLSX === 'undefined') {
        throw new Error('Excel library did not load. Check your internet connection and refresh the page.');
      }
      const data = await file.arrayBuffer();
      state.workbook = XLSX.read(data, { type: 'array', cellDates: false, raw: true });
      state.fileName = file.name;

      sheetSelect.innerHTML = '';
      state.workbook.SheetNames.forEach((name) => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        sheetSelect.appendChild(option);
      });

      fileMeta.textContent = `${file.name} • ${(file.size / 1024).toFixed(1)} KB`;
      fileMeta.classList.remove('hidden');
      sheetSection.classList.remove('hidden');
      settingsSection.classList.remove('hidden');
      parseSelectedSheet();
    } catch (error) {
      state.workbook = null;
      state.rows = [];
      showStatus(error.message || String(error), 'error');
    }
  }

  function parseSelectedSheet() {
    if (!state.workbook) return;
    try {
      const name = sheetSelect.value || state.workbook.SheetNames[0];
      const rows = normalizeSheet(state.workbook.Sheets[name]);
      state.rows = rows;
      updateDataPreview(rows);

      const duplicates = duplicateCoordinateGroups(rows);
      if (duplicates.length) {
        const text = duplicates.map((g) => `S.No ${g.labels.join(', ')} at ${g.key}`).join('; ');
        showStatus(
          `Loaded ${rows.length} valid site rows. Duplicate GPS coordinates detected: ${text}. ` +
          'Their arrow tips will share the same exact point while the numbered bubbles are separated.',
          'warning'
        );
      } else {
        showStatus(`Loaded ${rows.length} site rows with valid GPS coordinates.`, 'success');
      }
      renderInfo.textContent = 'Ready to generate.';
    } catch (error) {
      state.rows = [];
      showStatus(error.message || String(error), 'error');
      rowCount.textContent = '';
      dataPreview.classList.add('hidden');
    }
  }

  function validateCoordinate(lat, lon) {
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) throw new Error('Missing latitude/longitude.');
    if (lat < -90 || lat > 90) throw new Error(`Invalid latitude: ${lat}`);
    if (lon < -180 || lon > 180) throw new Error(`Invalid longitude: ${lon}`);
  }

  function latlonToNorm(lat, lon) {
    validateCoordinate(lat, lon);
    const clampedLat = Math.max(Math.min(lat, 85.05112878), -85.05112878);
    const x = (lon + 180) / 360;
    const sinLat = Math.sin((clampedLat * Math.PI) / 180);
    const y = 0.5 - Math.log((1 + sinLat) / (1 - sinLat)) / (4 * Math.PI);
    return [x, y];
  }

  function normToLatlon(x, y) {
    const lon = x * 360 - 180;
    const n = Math.PI - 2 * Math.PI * y;
    const lat = (Math.atan(Math.sinh(n)) * 180) / Math.PI;
    return [lat, lon];
  }

  function bboxNorm(points) {
    const xs = points.map((p) => p[0]);
    const ys = points.map((p) => p[1]);
    return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
  }

  function chooseZoom(points, safeWidth, safeHeight, minZoom = 4, maxZoom = 18) {
    const [minX, minY, maxX, maxY] = bboxNorm(points);
    const dx = Math.max(maxX - minX, 1e-9);
    const dy = Math.max(maxY - minY, 1e-9);
    for (let zoom = maxZoom; zoom >= minZoom; zoom -= 1) {
      const world = 256 * (2 ** zoom);
      if (dx * world <= safeWidth && dy * world <= safeHeight) return zoom;
    }
    return minZoom;
  }

  function chooseCenter(points, zoom, outWidth, outHeight, targetX, targetY) {
    const [minX, minY, maxX, maxY] = bboxNorm(points);
    const dataX = (minX + maxX) / 2;
    const dataY = (minY + maxY) / 2;
    const world = 256 * (2 ** zoom);
    let centerX = dataX + (outWidth / 2 - targetX) / world;
    let centerY = dataY + (outHeight / 2 - targetY) / world;
    centerX = ((centerX % 1) + 1) % 1;
    centerY = Math.max(0, Math.min(1, centerY));
    return [centerX, centerY];
  }

  function pointToPixel(pointNorm, centerNorm, zoom, width, height) {
    const world = 256 * (2 ** zoom);
    let dx = pointNorm[0] - centerNorm[0];
    if (dx > 0.5) dx -= 1;
    else if (dx < -0.5) dx += 1;
    const dy = pointNorm[1] - centerNorm[1];
    return [width / 2 + dx * world, height / 2 + dy * world];
  }

  function pixelToLatlon(x, y, centerNorm, zoom, width, height) {
    const world = 256 * (2 ** zoom);
    let nx = centerNorm[0] + (x - width / 2) / world;
    let ny = centerNorm[1] + (y - height / 2) / world;
    nx = ((nx % 1) + 1) % 1;
    ny = Math.max(0, Math.min(1, ny));
    return normToLatlon(nx, ny);
  }

  function verifyProjectionAccuracy(rows, pixels, centerNorm, zoom, width, height) {
    const failures = [];
    rows.forEach((row, i) => {
      const [lat2, lon2] = pixelToLatlon(pixels[i][0], pixels[i][1], centerNorm, zoom, width, height);
      if (Math.abs(row.Lat - lat2) > 1e-7 || Math.abs(row.Long - lon2) > 1e-7) {
        failures.push(`S.No ${formatCell(row['S.No'])}`);
      }
    });
    if (failures.length) throw new Error(`Internal GPS projection check failed for: ${failures.join(', ')}`);
  }

  function tableGeometry(width, tableWidthRatio = 0.49) {
    const tableW = Math.round(width * tableWidthRatio);
    const rightMargin = Math.round(width * 0.018);
    const x0 = width - rightMargin - tableW;
    return { x0, tableW, rightMargin };
  }

  function columnEdges(x0, tableW) {
    const weights = [0.075, 0.098, 0.105, 0.148, 0.460, 0.057, 0.057];
    const edges = [x0];
    let acc = x0;
    for (let i = 0; i < weights.length - 1; i += 1) {
      acc += Math.round(tableW * weights[i]);
      edges.push(acc);
    }
    edges.push(x0 + tableW);
    return edges;
  }

  function segmentIntersection(p1, p2, q1, q2) {
    const orient = (a, b, c) => (b[0] - a[0]) * (c[1] - a[1]) - (b[1] - a[1]) * (c[0] - a[0]);
    const onSegment = (a, b, c) => (
      Math.min(a[0], b[0]) - 1e-6 <= c[0] && c[0] <= Math.max(a[0], b[0]) + 1e-6 &&
      Math.min(a[1], b[1]) - 1e-6 <= c[1] && c[1] <= Math.max(a[1], b[1]) + 1e-6
    );

    const o1 = orient(p1, p2, q1);
    const o2 = orient(p1, p2, q2);
    const o3 = orient(q1, q2, p1);
    const o4 = orient(q1, q2, p2);

    if (Math.abs(o1) < 1e-6 && onSegment(p1, p2, q1)) return true;
    if (Math.abs(o2) < 1e-6 && onSegment(p1, p2, q2)) return true;
    if (Math.abs(o3) < 1e-6 && onSegment(q1, q2, p1)) return true;
    if (Math.abs(o4) < 1e-6 && onSegment(q1, q2, p2)) return true;
    return (o1 > 0) !== (o2 > 0) && (o3 > 0) !== (o4 > 0);
  }

  function pointSegmentDistance(point, a, b) {
    const [px, py] = point;
    const [ax, ay] = a;
    const [bx, by] = b;
    const dx = bx - ax;
    const dy = by - ay;
    if (Math.abs(dx) < 1e-9 && Math.abs(dy) < 1e-9) return Math.hypot(px - ax, py - ay);
    let t = ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy);
    t = Math.max(0, Math.min(1, t));
    const cx = ax + t * dx;
    const cy = ay + t * dy;
    return Math.hypot(px - cx, py - cy);
  }

  function calloutGeometry(anchorX, anchorY, bubbleX, bubbleY, radius) {
    const ax = Number(anchorX);
    const ay = Number(anchorY);
    const bx = Number(bubbleX);
    const by = Number(bubbleY);
    const r = Math.max(11, Math.round(radius));

    const lineW = Math.max(2, Math.round(r * 0.17));
    const haloW = lineW + Math.max(2, Math.round(r * 0.10));
    const arrowLen = Math.max(8, r * 0.48);
    const arrowHalf = Math.max(4.5, r * 0.24);

    const dx = ax - bx;
    const dy = ay - by;
    const dist = Math.max(1e-6, Math.hypot(dx, dy));
    const ux = dx / dist;
    const uy = dy / dist;
    const startX = bx + ux * r;
    const startY = by + uy * r;

    const hdir = ax > bx ? 1 : -1;
    const elbowLen = Math.min(Math.max(r * 0.42, 6), r * 0.78);
    const elbowX = startX + hdir * elbowLen;
    const elbowY = startY;

    const vx = ax - elbowX;
    const vy = ay - elbowY;
    const vdist = Math.max(1e-6, Math.hypot(vx, vy));
    const vux = vx / vdist;
    const vuy = vy / vdist;
    const baseX = ax - vux * arrowLen;
    const baseY = ay - vuy * arrowLen;
    const px = -vuy;
    const py = vux;

    return {
      lineW,
      haloW,
      start: [startX, startY],
      elbow: [elbowX, elbowY],
      base: [baseX, baseY],
      tip: [ax, ay],
      arrowPts: [
        [ax, ay],
        [baseX + px * arrowHalf, baseY + py * arrowHalf],
        [baseX - px * arrowHalf, baseY - py * arrowHalf],
      ],
      haloArrowPts: [
        [ax, ay],
        [baseX + px * arrowHalf * 1.16, baseY + py * arrowHalf * 1.16],
        [baseX - px * arrowHalf * 1.16, baseY - py * arrowHalf * 1.16],
      ],
      segments: [
        [[startX, startY], [elbowX, elbowY]],
        [[elbowX, elbowY], [baseX, baseY]],
      ],
      bubble: [bx, by],
      radius: r,
    };
  }

  function placeCalloutBubbles(anchors, width, height, mapRight, bubbleRadius) {
    const r = Number(bubbleRadius);
    const margin = Math.max(12, Math.round(height * 0.012));
    const gap = Math.max(8, Math.round(height * 0.009));
    const maxRight = mapRight - margin - r;
    const minX = margin + r;
    const minY = margin + r;
    const maxY = height - margin - r;

    const results = Array(anchors.length).fill(null);
    const placedBubbles = [];
    const placedGeoms = [];

    const d1 = Math.max(r * 2.8, height * 0.048);
    const d2 = Math.max(r * 4.1, height * 0.072);
    const d3 = Math.max(r * 5.5, height * 0.097);
    const d4 = Math.max(r * 7.0, height * 0.125);
    const dy1 = Math.max(r * 1.8, height * 0.030);
    const dy2 = Math.max(r * 3.2, height * 0.052);
    const dy3 = Math.max(r * 4.7, height * 0.076);

    const order = anchors
      .map((p, i) => ({ i, x: p[0], y: p[1] }))
      .sort((a, b) => a.y - b.y || a.x - b.x || a.i - b.i)
      .map((item) => item.i);

    order.forEach((idx) => {
      const [ax, ay] = anchors[idx];
      const preferRight = ax + d2 <= maxRight;
      const primaryDir = preferRight ? 1 : -1;
      const directions = [primaryDir, -primaryDir];
      const candidates = [];
      let rank = 0;

      directions.forEach((direction) => {
        [d1, d2, d3, d4].forEach((dx) => {
          [0, -dy1, dy1, -dy2, dy2, -dy3, dy3].forEach((yoff) => {
            candidates.push([ax + direction * dx, ay + yoff, rank]);
            rank += 0.18;
          });
        });
      });

      directions.forEach((direction) => {
        [
          [d1, -dy1], [d1, dy1], [d2, -dy2], [d2, dy2], [d3, -dy1], [d3, dy1],
        ].forEach(([dx, yoff]) => {
          candidates.push([ax + direction * dx, ay + yoff, rank]);
          rank += 0.16;
        });
      });

      let best = null;
      let bestScore = Number.POSITIVE_INFINITY;

      candidates.forEach(([cx0, cy0, rankScore]) => {
        const cx = Math.min(Math.max(cx0, minX), maxRight);
        const cy = Math.min(Math.max(cy0, minY), maxY);
        const leaderLen = Math.hypot(cx - ax, cy - ay);
        let score = rankScore * 10;

        const preferredLen = Math.max(height * 0.075, r * 4.3);
        const hardLen = Math.max(height * 0.175, r * 9.0);
        score += leaderLen * 0.42;
        if (leaderLen > preferredLen) score += (leaderLen - preferredLen) * 5.5;
        if (leaderLen > hardLen) score += (leaderLen - hardLen) * 120 + 100000;

        placedBubbles.forEach(([obx, oby]) => {
          const dist = Math.hypot(cx - obx, cy - oby);
          const required = r * 2 + gap;
          if (dist < required) score += 120000 + (required - dist) * 1800;
        });

        anchors.forEach(([px, py], j) => {
          if (j === idx) return;
          const dist = Math.hypot(cx - px, cy - py);
          const required = r + gap * 1.25;
          if (dist < required) score += 45000 + (required - dist) * 900;
        });

        const geom = calloutGeometry(ax, ay, cx, cy, r);

        placedGeoms.forEach((existing) => {
          geom.segments.forEach((seg1) => {
            existing.segments.forEach((seg2) => {
              if (segmentIntersection(seg1[0], seg1[1], seg2[0], seg2[1])) score += 150000;
            });
          });

          const [ebx, eby] = existing.bubble;
          const erad = existing.radius;
          geom.segments.forEach(([segA, segB]) => {
            const d = pointSegmentDistance([ebx, eby], segA, segB);
            if (d < erad + gap * 0.45) score += 30000 + (erad + gap * 0.45 - d) * 1000;
          });

          existing.segments.forEach(([segA, segB]) => {
            const d = pointSegmentDistance([cx, cy], segA, segB);
            if (d < r + gap * 0.45) score += 50000 + (r + gap * 0.45 - d) * 1000;
          });
        });

        anchors.forEach((pt, j) => {
          if (j === idx) return;
          geom.segments.forEach(([segA, segB]) => {
            const d = pointSegmentDistance(pt, segA, segB);
            if (d < gap * 0.65) score += 16000 + (gap * 0.65 - d) * 650;
          });
        });

        if (score < bestScore) {
          bestScore = score;
          best = [cx, cy, geom];
        }
      });

      if (!best) {
        const direction = ax + d1 <= maxRight ? 1 : -1;
        const cx = Math.min(Math.max(ax + direction * d1, minX), maxRight);
        const cy = Math.min(Math.max(ay, minY), maxY);
        best = [cx, cy, calloutGeometry(ax, ay, cx, cy, r)];
      }

      const [cx, cy, geom] = best;
      results[idx] = [cx, cy];
      placedBubbles.push([cx, cy]);
      placedGeoms.push(geom);
    });

    return results.map((p, i) => p || anchors[i]);
  }

  function drawPolyline(context, points, color, width) {
    context.beginPath();
    context.moveTo(points[0][0], points[0][1]);
    for (let i = 1; i < points.length; i += 1) context.lineTo(points[i][0], points[i][1]);
    context.strokeStyle = color;
    context.lineWidth = width;
    context.lineCap = 'round';
    context.lineJoin = 'round';
    context.stroke();
  }

  function fillPolygon(context, points, color) {
    context.beginPath();
    context.moveTo(points[0][0], points[0][1]);
    for (let i = 1; i < points.length; i += 1) context.lineTo(points[i][0], points[i][1]);
    context.closePath();
    context.fillStyle = color;
    context.fill();
  }

  function drawCalloutMarker(context, anchorX, anchorY, bubbleX, bubbleY, label, color, radius) {
    const geom = calloutGeometry(anchorX, anchorY, bubbleX, bubbleY, radius);
    const [startX, startY] = geom.start;
    const [elbowX, elbowY] = geom.elbow;
    const [baseX, baseY] = geom.base;
    const [ax, ay] = geom.tip;

    drawPolyline(context, [[startX, startY], [elbowX, elbowY], [baseX, baseY]], 'rgba(255,255,255,.94)', geom.haloW);
    drawPolyline(context, [[startX, startY], [elbowX, elbowY], [baseX, baseY]], color, geom.lineW);
    fillPolygon(context, geom.haloArrowPts, 'rgba(255,255,255,.94)');
    fillPolygon(context, geom.arrowPts, color);

    const r = geom.radius;
    const anchorOuter = Math.max(3, Math.round(r * 0.18));
    const anchorInner = Math.max(2, Math.round(r * 0.12));
    context.beginPath();
    context.arc(ax, ay, anchorOuter, 0, Math.PI * 2);
    context.fillStyle = WHITE;
    context.fill();
    context.beginPath();
    context.arc(ax, ay, anchorInner, 0, Math.PI * 2);
    context.fillStyle = color;
    context.fill();

    const [bx, by] = geom.bubble;
    const shadow = Math.max(2, Math.round(r * 0.08));
    context.beginPath();
    context.arc(bx + shadow, by + shadow, r, 0, Math.PI * 2);
    context.fillStyle = 'rgba(0,0,0,.18)';
    context.fill();
    context.beginPath();
    context.arc(bx, by, r, 0, Math.PI * 2);
    context.fillStyle = color;
    context.fill();

    const len = String(label).length;
    const fontScale = len <= 1 ? 1.02 : len === 2 ? 0.86 : 0.70;
    const fontSize = Math.max(8, Math.round(r * fontScale));
    context.font = `700 ${fontSize}px Arial, Helvetica, sans-serif`;
    context.fillStyle = WHITE;
    context.textAlign = 'center';
    context.textBaseline = 'middle';
    context.fillText(String(label), bx, by + 0.5);
  }

  function splitLongWord(context, word, maxWidth) {
    if (!word) return [''];
    if (context.measureText(word).width <= maxWidth) return [word];
    const parts = [];
    let current = '';
    for (const ch of word) {
      const trial = current + ch;
      if (current && context.measureText(trial).width > maxWidth) {
        parts.push(current);
        current = ch;
      } else current = trial;
    }
    if (current) parts.push(current);
    return parts;
  }

  function wrapTextPx(context, value, maxWidth) {
    const text = String(value ?? '').trim();
    if (!text) return [''];
    const tokens = [];
    text.split(/\s+/).forEach((word) => tokens.push(...splitLongWord(context, word, maxWidth)));
    const lines = [];
    let current = '';
    tokens.forEach((token) => {
      const trial = current ? `${current} ${token}` : token;
      if (context.measureText(trial).width <= maxWidth) current = trial;
      else {
        if (current) lines.push(current);
        current = token;
      }
    });
    if (current) lines.push(current);
    return lines;
  }

  function roundRectPath(context, x, y, w, h, radius) {
    const r = Math.min(radius, w / 2, h / 2);
    context.beginPath();
    context.moveTo(x + r, y);
    context.lineTo(x + w - r, y);
    context.quadraticCurveTo(x + w, y, x + w, y + r);
    context.lineTo(x + w, y + h - r);
    context.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
    context.lineTo(x + r, y + h);
    context.quadraticCurveTo(x, y + h, x, y + h - r);
    context.lineTo(x, y + r);
    context.quadraticCurveTo(x, y, x + r, y);
    context.closePath();
  }

  function measureTableRows(context, rows, edges, fontSize, minRowH, padY) {
    const lineH = Math.max(11, fontSize * 1.20);
    const lineGap = Math.max(2, lineH * 0.16);
    const rowHeights = [];
    const wrappedRows = [];

    rows.forEach((row) => {
      const wrapped = {};
      let maxCount = 1;
      TABLE_COLUMNS.slice(1).forEach((col, ci) => {
        const c = ci + 1;
        const value = formatCell(row[col]);
        context.font = `${col === 'W' || col === 'H' ? '700' : '400'} ${fontSize}px Arial, Helvetica, sans-serif`;
        const maxWidth = Math.max(10, edges[c + 1] - edges[c] - 16);
        const lines = wrapTextPx(context, value, maxWidth);
        wrapped[col] = lines;
        maxCount = Math.max(maxCount, lines.length);
      });
      const contentH = maxCount * lineH + Math.max(0, maxCount - 1) * lineGap + padY * 2;
      rowHeights.push(Math.max(minRowH, contentH));
      wrappedRows.push(wrapped);
    });

    return { rowHeights, wrappedRows, lineH, lineGap };
  }

  function drawCellLines(context, lines, box, fontSize, weight, color, align = 'center') {
    const [x0, y0, x1, y1] = box;
    const lineH = Math.max(11, fontSize * 1.20);
    const lineGap = Math.max(2, lineH * 0.16);
    const totalH = lines.length * lineH + Math.max(0, lines.length - 1) * lineGap;
    let y = y0 + (y1 - y0 - totalH) / 2 + lineH / 2;
    context.font = `${weight} ${fontSize}px Arial, Helvetica, sans-serif`;
    context.fillStyle = color;
    context.textBaseline = 'middle';
    context.textAlign = align === 'left' ? 'left' : 'center';
    const x = align === 'left' ? x0 + 8 : (x0 + x1) / 2;
    lines.forEach((line) => {
      context.fillText(line, x, y);
      y += lineH + lineGap;
    });
  }

  function drawTable(context, rows, highlightLabels, width, height, tableWidthRatio = 0.49) {
    const { x0, tableW } = tableGeometry(width, tableWidthRatio);
    const edges = columnEdges(x0, tableW);
    const n = rows.length;
    const maxTotalH = height * 0.92;
    const headerH = Math.max(44, height * 0.055);
    const availableBodyH = maxTotalH - headerH;
    const minRowH = Math.max(30, Math.min(height * 0.047, availableBodyH / Math.max(n, 1)));
    const padY = Math.max(4, height * 0.0045);
    const startFont = Math.max(12, Math.floor(height * 0.0155));
    const minFont = Math.max(9, Math.floor(height * 0.009));

    let chosen = null;
    for (let fontSize = startFont; fontSize >= minFont; fontSize -= 1) {
      const measured = measureTableRows(context, rows, edges, fontSize, minRowH, padY);
      const totalH = headerH + measured.rowHeights.reduce((a, b) => a + b, 0);
      if (totalH <= maxTotalH) {
        chosen = { fontSize, totalH, ...measured };
        break;
      }
    }

    if (!chosen) {
      throw new Error(
        `${n} site rows cannot fit with full readable addresses at ${width}x${height}. ` +
        'Choose a larger output size or split the Excel list into fewer rows.'
      );
    }

    const y0 = Math.max(height * 0.035, (height - chosen.totalH) / 2);
    const radius = Math.max(10, height * 0.011);
    const shadow = Math.max(6, height * 0.007);

    roundRectPath(context, x0 + shadow, y0 + shadow, tableW, chosen.totalH, radius);
    context.fillStyle = 'rgba(0,0,0,.18)';
    context.fill();

    roundRectPath(context, x0, y0, tableW, chosen.totalH, radius);
    context.fillStyle = 'rgba(255,255,255,.98)';
    context.fill();
    context.lineWidth = 2;
    context.strokeStyle = '#B4B9BE';
    context.stroke();

    roundRectPath(context, x0, y0, tableW, headerH + 8, radius);
    context.fillStyle = NAVY;
    context.fill();
    context.fillRect(x0, y0 + headerH - 8, tableW, 8);

    context.strokeStyle = GRID;
    context.lineWidth = 1;
    for (let i = 1; i < edges.length - 1; i += 1) {
      context.beginPath();
      context.moveTo(edges[i], y0);
      context.lineTo(edges[i], y0 + chosen.totalH);
      context.stroke();
    }

    const headerFontSize = Math.max(11, Math.round(headerH * 0.34));
    TABLE_COLUMNS.forEach((col, i) => {
      drawCellLines(context, [col], [edges[i], y0, edges[i + 1], y0 + headerH], headerFontSize, 700, WHITE, 'center');
    });

    let top = y0 + headerH;
    rows.forEach((row, rIndex) => {
      const rowH = chosen.rowHeights[rIndex];
      const bottom = top + rowH;
      context.strokeStyle = GRID;
      context.beginPath();
      context.moveTo(x0, top);
      context.lineTo(x0 + tableW, top);
      context.stroke();

      const label = formatCell(row['S.No']);
      const cx = (edges[0] + edges[1]) / 2;
      const cy = (top + bottom) / 2;
      const badgeRadius = Math.max(11, Math.min(19, rowH * 0.25));
      const badgeColor = highlightLabels.has(label) ? RED : NAVY;
      context.beginPath();
      context.arc(cx, cy, badgeRadius, 0, Math.PI * 2);
      context.fillStyle = badgeColor;
      context.fill();
      context.font = `700 ${Math.max(9, badgeRadius * 1.03)}px Arial, Helvetica, sans-serif`;
      context.fillStyle = WHITE;
      context.textAlign = 'center';
      context.textBaseline = 'middle';
      context.fillText(label, cx, cy + 0.5);

      TABLE_COLUMNS.slice(1).forEach((col, ci) => {
        const c = ci + 1;
        const lines = chosen.wrappedRows[rIndex][col];
        drawCellLines(
          context,
          lines,
          [edges[c], top, edges[c + 1], bottom],
          chosen.fontSize,
          col === 'W' || col === 'H' ? 700 : 400,
          TEXT,
          col === 'Location' ? 'left' : 'center'
        );
      });
      top = bottom;
    });

    context.strokeStyle = GRID;
    context.beginPath();
    context.moveTo(x0, y0 + chosen.totalH);
    context.lineTo(x0 + tableW, y0 + chosen.totalH);
    context.stroke();
  }

  function drawOsmAttribution(context, width, height) {
    const text = 'Map data © OpenStreetMap contributors';
    const fontSize = Math.max(11, Math.round(height * 0.0105));
    context.font = `400 ${fontSize}px Arial, Helvetica, sans-serif`;
    const tw = context.measureText(text).width;
    const pad = 6;
    const x = width - tw - pad * 2 - 8;
    const h = fontSize + pad * 2 + 2;
    const y = height - h - 5;
    context.fillStyle = 'rgba(255,255,255,.88)';
    context.fillRect(x, y, tw + pad * 2, h);
    context.fillStyle = '#333333';
    context.textAlign = 'left';
    context.textBaseline = 'middle';
    context.fillText(text, x + pad, y + h / 2 + 0.5);
  }

  function parseHighlights(text) {
    return new Set(String(text || '').split(/[,;\s]+/).map((s) => s.trim()).filter(Boolean));
  }

  function mod(value, n) {
    return ((value % n) + n) % n;
  }

  async function fetchTileBitmap(z, x, y) {
    const n = 2 ** z;
    const wrappedX = mod(x, n);
    if (y < 0 || y >= n) return null;
    const key = `${z}/${wrappedX}/${y}`;
    if (state.tileCache.has(key)) return state.tileCache.get(key);

    const promise = (async () => {
      const url = `https://tile.openstreetmap.org/${z}/${wrappedX}/${y}.png`;
      const response = await fetch(url, { mode: 'cors', cache: 'force-cache' });
      if (!response.ok) throw new Error(`Map tile failed to load (${response.status}).`);
      const blob = await response.blob();
      if ('createImageBitmap' in window) return createImageBitmap(blob);
      return new Promise((resolve, reject) => {
        const img = new Image();
        img.crossOrigin = 'anonymous';
        img.onload = () => resolve(img);
        img.onerror = () => reject(new Error('A map tile could not be decoded.'));
        img.src = URL.createObjectURL(blob);
      });
    })();

    state.tileCache.set(key, promise);
    try {
      return await promise;
    } catch (error) {
      state.tileCache.delete(key);
      throw error;
    }
  }

  async function drawOsmBase(context, centerNorm, zoom, width, height, onProgress) {
    context.fillStyle = '#EAEAEA';
    context.fillRect(0, 0, width, height);

    const world = 256 * (2 ** zoom);
    const centerPxX = centerNorm[0] * world;
    const centerPxY = centerNorm[1] * world;
    const left = centerPxX - width / 2;
    const top = centerPxY - height / 2;
    const minTx = Math.floor(left / 256);
    const maxTx = Math.floor((left + width) / 256);
    const minTy = Math.floor(top / 256);
    const maxTy = Math.floor((top + height) / 256);

    const tasks = [];
    for (let ty = minTy; ty <= maxTy; ty += 1) {
      for (let tx = minTx; tx <= maxTx; tx += 1) {
        tasks.push({ tx, ty });
      }
    }

    let done = 0;
    const totalCount = tasks.length;
    const workers = Array.from({ length: Math.min(8, tasks.length) }, async () => {
      while (tasks.length) {
        const item = tasks.shift();
        if (!item) break;
        const tile = await fetchTileBitmap(zoom, item.tx, item.ty);
        if (tile) {
          const px = Math.round(item.tx * 256 - left);
          const py = Math.round(item.ty * 256 - top);
          context.drawImage(tile, px, py, 256, 256);
        }
        done += 1;
        if (onProgress) onProgress(done, totalCount);
      }
    });
    await Promise.all(workers);
  }

  async function renderMap(rows, width, height, highlightLabels) {
    canvas.width = width;
    canvas.height = height;
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, width, height);

    const pointNorms = rows.map((row) => latlonToNorm(row.Lat, row.Long));
    const safeW = width * 0.48;
    const safeH = height * 0.76;
    const zoom = chooseZoom(pointNorms, safeW, safeH);
    const center = chooseCenter(pointNorms, zoom, width, height, width * 0.295, height * 0.52);

    await drawOsmBase(ctx, center, zoom, width, height, (done, total) => {
      renderInfo.textContent = `Loading map tiles ${done}/${total}…`;
    });

    const markerRadius = Math.max(13, Math.round(height * 0.0165));
    const markerPixels = pointNorms.map((p) => pointToPixel(p, center, zoom, width, height));
    verifyProjectionAccuracy(rows, markerPixels, center, zoom, width, height);

    const { x0: tableX0 } = tableGeometry(width);
    const mapRight = tableX0 - Math.max(10, Math.round(width * 0.012));
    const bubblePixels = placeCalloutBubbles(markerPixels, width, height, mapRight, markerRadius);

    rows.forEach((row, i) => {
      const label = formatCell(row['S.No']);
      const color = highlightLabels.has(label) ? RED : BLACK;
      drawCalloutMarker(
        ctx,
        markerPixels[i][0],
        markerPixels[i][1],
        bubblePixels[i][0],
        bubblePixels[i][1],
        label,
        color,
        markerRadius
      );
    });

    drawTable(ctx, rows, highlightLabels, width, height);
    drawOsmAttribution(ctx, width, height);
    ctx.restore();
  }

  async function generate() {
    if (!state.rows.length) {
      showStatus('Upload a valid Excel file first.', 'error');
      return;
    }

    const [width, height] = sizeSelect.value.split('x').map(Number);
    const highlights = parseHighlights(extraHighlights.value);
    if (highlightLast.checked && state.rows.length) {
      highlights.add(formatCell(state.rows[state.rows.length - 1]['S.No']));
    }

    generateBtn.disabled = true;
    generateBtn.textContent = 'Generating…';
    downloadBtn.classList.add('hidden');
    showStatus('Generating map. GPS anchors are being projection-checked before drawing.', 'info');
    renderInfo.textContent = 'Preparing map…';

    try {
      await renderMap(state.rows, width, height, highlights);
      emptyState.classList.add('hidden');
      canvas.classList.remove('hidden');
      canvasStage.classList.remove('empty');

      state.lastBlob = await new Promise((resolve, reject) => {
        canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error('PNG export failed.')), 'image/png', 1);
      });
      state.generated = true;
      downloadBtn.classList.remove('hidden');
      renderInfo.textContent = `${width} × ${height} • Exact GPS callout map`;
      showStatus(`Map generated successfully at ${width} × ${height}.`, 'success');
    } catch (error) {
      console.error(error);
      showStatus(
        `${error.message || String(error)}\n\nIf map tiles fail, confirm the browser has internet access and that OpenStreetMap tile requests are not being blocked by your network.`,
        'error'
      );
      renderInfo.textContent = 'Generation failed.';
    } finally {
      generateBtn.disabled = false;
      generateBtn.textContent = 'Generate exact GPS map';
    }
  }

  function downloadPng() {
    if (!state.lastBlob) return;
    const url = URL.createObjectURL(state.lastBlob);
    const a = document.createElement('a');
    const base = state.fileName ? state.fileName.replace(/\.[^.]+$/, '') : 'site_location_map';
    a.href = url;
    a.download = `${base}_site_location_map.png`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  fileInput.addEventListener('change', (event) => handleFile(event.target.files?.[0]));
  sheetSelect.addEventListener('change', parseSelectedSheet);
  generateBtn.addEventListener('click', generate);
  downloadBtn.addEventListener('click', downloadPng);

  ['dragenter', 'dragover'].forEach((eventName) => {
    dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropZone.classList.add('dragging');
    });
  });
  ['dragleave', 'drop'].forEach((eventName) => {
    dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropZone.classList.remove('dragging');
    });
  });
  dropZone.addEventListener('drop', (event) => {
    const file = event.dataTransfer?.files?.[0];
    if (file) handleFile(file);
  });
})();
