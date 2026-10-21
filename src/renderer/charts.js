'use strict';

// Draws a chart for the preview from the helper's description of it (see describe() in
// py/charts.py), laid out the way the converted Excel chart is: the title at the top, the plot
// where the chart's body was in Numbers, and the legend where Numbers put it. Built with DOM calls
// only, so text from the file can never become markup.

const SVG_NS = 'http://www.w3.org/2000/svg';
const CHART_PT = 4 / 3; // chart fonts and line widths are in points

function svgNode(tag, attrs = {}, text) {
  const node = document.createElementNS(SVG_NS, tag);
  for (const [key, value] of Object.entries(attrs)) {
    if (value !== undefined && value !== null && value !== false) node.setAttribute(key, value);
  }
  if (text !== undefined) node.textContent = text;
  return node;
}

function fontAttrs(style, fallbackSize = 10) {
  const s = style || {};
  return {
    'font-family': `"${s.font || 'Arial'}", Arial, sans-serif`,
    'font-size': ((s.size || fallbackSize) * CHART_PT).toFixed(1),
    'font-weight': s.bold ? 700 : 400,
    'font-style': s.italic ? 'italic' : null,
    fill: s.color || '#000000',
  };
}

function fontPx(style, fallbackSize = 10) {
  return ((style && style.size) || fallbackSize) * CHART_PT;
}

function textWidth(text, style, fallbackSize) {
  return String(text).length * fontPx(style, fallbackSize) * 0.5; // an average character in Arial
}

// Axis steps of 1, 2 or 5 times a power of ten, from zero unless the values sit well away from
// it, much as Excel picks them.
function niceScale(lo, hi, ticks = 5) {
  if (!Number.isFinite(lo) || !Number.isFinite(hi)) { lo = 0; hi = 1; }
  if (lo > 0 && (hi - lo) >= hi / 6) lo = 0;
  if (hi < 0 && (hi - lo) >= -lo / 6) hi = 0;
  if (lo === hi) { hi = lo === 0 ? 1 : lo + Math.abs(lo) * 0.1; }
  const raw = (hi - lo) / ticks;
  const magnitude = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 5, 10].map((m) => m * magnitude).find((s) => s >= raw * 0.999);
  return { lo: Math.floor(lo / step + 1e-9) * step, hi: Math.ceil(hi / step - 1e-9) * step, step };
}

function formatValue(value, format, step) {
  const f = format || {};
  let v = value;
  if (f.percent) v *= 100;
  let places = f.decimals;
  if (places === undefined || places === null) {
    places = step && step < 1 ? Math.min(4, Math.ceil(-Math.log10(step * (f.percent ? 100 : 1)) - 1e-9)) : 0;
    if (!step) places = Number.isInteger(v) ? 0 : Math.min(2, (String(v).split('.')[1] || '').length);
  }
  const text = Math.abs(v).toLocaleString('en-US', {
    minimumFractionDigits: places, maximumFractionDigits: places, useGrouping: Boolean(f.thousands),
  });
  return `${v < 0 ? '-' : ''}${f.prefix || ''}${text}${f.suffix || ''}`;
}

function seriesColour(s, i) {
  return s.color || s.line || (s.points && s.points[i]) || '#4472C4';
}

function drawChart(c) {
  const W = c.w;
  const H = c.h;
  const root = svgNode('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, class: 'chart' });
  const plot = { x: c.plot[0] * W, y: c.plot[1] * H, w: c.plot[2] * W, h: c.plot[3] * H };

  if (c.kind === 'pie' || c.kind === 'doughnut') drawPie(root, c, plot);
  else if (c.kind === 'radar') drawRadar(root, c, plot);
  else drawAxes(root, c, plot);

  if (c.title) {
    root.append(svgNode('text', { x: W / 2, y: fontPx(c.titleStyle, 12) * 1.25, 'text-anchor': 'middle',
      ...fontAttrs(c.titleStyle, 12) }, c.title));
  }
  if (c.legend) drawLegend(root, c, { x: c.legend[0] * W, y: c.legend[1] * H, w: c.legend[2] * W, h: c.legend[3] * H });
  return root;
}

function drawAxes(root, c, P) {
  const scatter = c.kind === 'scatter';
  const horizontal = c.kind === 'bar';
  const stacked = c.stacked;
  const count = Math.max(c.categories.length, ...c.series.map((s) => s.values.length), 1);
  const primary = c.series.filter((s) => !s.y2);
  const secondary = c.series.filter((s) => s.y2);

  const range = (list) => {
    let lo = Infinity;
    let hi = -Infinity;
    if (stacked) {
      for (let i = 0; i < count; i++) {
        let up = 0;
        let down = 0;
        for (const s of list) { const v = s.values[i] || 0; if (v >= 0) up += v; else down += v; }
        hi = Math.max(hi, up); lo = Math.min(lo, down);
      }
    } else {
      for (const s of list) for (const v of s.values) if (v !== null) { lo = Math.min(lo, v); hi = Math.max(hi, v); }
    }
    return niceScale(Math.min(lo, hi), hi);
  };
  const y1 = range(primary.length ? primary : c.series);
  const y2 = secondary.length ? range(secondary) : null;
  const axisStyle = c.axisStyle;
  const labelGap = 6;

  // Where a value sits along the value axis, and where category i (or an x value) sits along the other.
  const valuePos = (v, sc) => {
    const t = (v - sc.lo) / (sc.hi - sc.lo);
    return horizontal ? P.x + t * P.w : P.y + P.h - t * P.h;
  };
  let xs = null;
  if (scatter) {
    const all = c.series.flatMap((s) => (s.x || []).filter((v) => v !== null));
    xs = niceScale(Math.min(...all), Math.max(...all));
  }
  const band = (horizontal ? P.h : P.w) / count;
  const categoryPos = (i) => (horizontal ? P.y + band * (i + 0.5) : P.x + band * (i + 0.5));
  const xPos = (v) => P.x + ((v - xs.lo) / (xs.hi - xs.lo)) * P.w;

  // Gridlines and value labels.
  const grid = svgNode('g', { stroke: c.gridlines || 'none', 'stroke-width': 0.75 });
  root.append(grid);
  for (let v = y1.lo; v <= y1.hi + y1.step / 2; v += y1.step) {
    const p = valuePos(v, y1);
    if (c.gridlines) {
      grid.append(horizontal ? svgNode('line', { x1: p, x2: p, y1: P.y, y2: P.y + P.h })
        : svgNode('line', { x1: P.x, x2: P.x + P.w, y1: p, y2: p }));
    }
    if (c.valueLabels !== false) {
      const label = formatValue(v, c.format, y1.step);
      root.append(horizontal
        ? svgNode('text', { x: p, y: P.y + P.h + fontPx(axisStyle) + labelGap, 'text-anchor': 'middle', ...fontAttrs(axisStyle) }, label)
        : svgNode('text', { x: P.x - labelGap, y: p + fontPx(axisStyle) * 0.35, 'text-anchor': 'end', ...fontAttrs(axisStyle) }, label));
    }
  }
  if (y2) {
    for (let v = y2.lo; v <= y2.hi + y2.step / 2; v += y2.step) {
      root.append(svgNode('text', { x: P.x + P.w + labelGap, y: valuePos(v, y2) + fontPx(axisStyle) * 0.35,
        ...fontAttrs(axisStyle) }, formatValue(v, null, y2.step)));
    }
  }

  // Bars and columns: clustered side by side, or stacked.
  const bars = c.series.filter((s) => s.kind === 'column' || s.kind === 'bar');
  if (bars.length) {
    const gap = (c.gap ?? 150) / 100;
    const overlap = stacked ? 1 : (c.overlap ?? 0) / 100;
    const k = stacked ? 1 : bars.length;
    const width = band / (1 + (k - 1) * (1 - overlap) + gap);
    const ups = new Array(count).fill(0);
    const downs = new Array(count).fill(0);
    bars.forEach((s, si) => {
      const sc = s.y2 && y2 ? y2 : y1;
      s.values.forEach((v, i) => {
        if (v === null) return;
        let from = 0;
        if (stacked) { from = v >= 0 ? ups[i] : downs[i]; if (v >= 0) ups[i] += v; else downs[i] += v; }
        const base = Math.min(Math.max(from, sc.lo), sc.hi);
        const end = Math.min(Math.max(from + v, sc.lo), sc.hi);
        const start = (horizontal ? P.y : P.x) + band * i + (gap * width) / 2 + (stacked ? 0 : si * width * (1 - overlap));
        const a = valuePos(base, sc);
        const b = valuePos(end, sc);
        const fill = (s.points && s.points[i]) || s.color || '#4472C4';
        root.append(horizontal
          ? svgNode('rect', { x: Math.min(a, b), y: start, width: Math.abs(b - a), height: width, fill })
          : svgNode('rect', { x: start, y: Math.min(a, b), width, height: Math.abs(b - a), fill }));
        if (s.labels) {
          const style = s.labelStyle || {};
          const size = fontPx(style, 12);
          const label = formatValue(v, c.format);
          const inside = s.labelPosition !== 'outside_end';
          if (horizontal) {
            const x = inside ? b - (v >= 0 ? 4 : -4) : b + (v >= 0 ? 4 : -4);
            root.append(svgNode('text', { x, y: start + width / 2 + size * 0.35,
              'text-anchor': (v >= 0) === inside ? 'end' : 'start', ...fontAttrs(style, 12) }, label));
          } else {
            const y = inside ? b + (v >= 0 ? size + 2 : -4) : b - (v >= 0 ? 4 : -size - 2);
            root.append(svgNode('text', { x: start + width / 2, y, 'text-anchor': 'middle', ...fontAttrs(style, 12) }, label));
          }
        }
      });
    });
  }

  // Areas, then lines and points on top.
  const cumulative = new Array(count).fill(0);
  for (const s of c.series.filter((x) => x.kind === 'area')) {
    const sc = s.y2 && y2 ? y2 : y1;
    const tops = s.values.map((v, i) => (stacked ? cumulative[i] : 0) + (v || 0));
    const bottoms = s.values.map((_, i) => (stacked ? cumulative[i] : 0));
    if (stacked) tops.forEach((t, i) => { cumulative[i] = t; });
    const edge = (vals) => vals.map((v, i) => `${categoryPos(i)},${valuePos(Math.min(Math.max(v, sc.lo), sc.hi), sc)}`);
    const points = [...edge(tops), ...edge(bottoms).reverse()].join(' ');
    root.append(svgNode('polygon', { points, fill: s.color || '#4472C4', 'fill-opacity': 0.85,
      stroke: s.line || 'none', 'stroke-width': (s.lineWidth || 1) * CHART_PT }));
  }
  for (const s of c.series.filter((x) => x.kind === 'line' || x.kind === 'scatter')) {
    const sc = s.y2 && y2 ? y2 : y1;
    const at = s.values.map((v, i) => {
      if (v === null) return null;
      const x = scatter ? (s.x && s.x[i] !== null ? xPos(s.x[i]) : null) : categoryPos(i);
      return x === null ? null : [x, valuePos(v, sc)];
    });
    const colour = s.line || s.color || '#4472C4';
    if (!scatter || s.joined) {
      root.append(svgNode('polyline', { points: at.filter(Boolean).map((p) => p.join(',')).join(' '),
        fill: 'none', stroke: colour, 'stroke-width': Math.max(1, (s.lineWidth || 2) * CHART_PT), 'stroke-linejoin': 'round' }));
    }
    if (s.marker) {
      const r = Math.max(2, (s.marker.size * CHART_PT) / 2);
      for (const p of at.filter(Boolean)) {
        root.append(svgNode('circle', { cx: p[0], cy: p[1], r, fill: s.marker.fill || colour,
          stroke: s.marker.outline || colour, 'stroke-width': 1.5 }));
      }
    }
  }

  // Axis titles: up the left (and right, for a second value axis) and along the bottom.
  const titleStyle = c.titleFont || axisStyle;
  const titleSize = fontPx(titleStyle);
  const sideTitle = (text, x) => root.append(svgNode('text', { x, y: P.y + P.h / 2, 'text-anchor': 'middle',
    transform: `rotate(-90 ${x} ${P.y + P.h / 2})`, ...fontAttrs(titleStyle) }, text));
  const leftTitle = horizontal ? c.categoryTitle : c.valueTitle;
  const bottomTitle = horizontal ? c.valueTitle : c.categoryTitle;
  if (leftTitle) sideTitle(leftTitle, titleSize);
  if (c.y2Title) sideTitle(c.y2Title, root.getAttribute('width') - titleSize * 0.4);
  if (bottomTitle) {
    root.append(svgNode('text', { x: P.x + P.w / 2, y: P.y + P.h + fontPx(axisStyle) * 2 + titleSize + labelGap,
      'text-anchor': 'middle', ...fontAttrs(titleStyle) }, bottomTitle));
  }

  // The category axis: a line along the bottom (or left, for bars) and its labels.
  if (c.categoryLine !== false) {
    root.append(horizontal
      ? svgNode('line', { x1: P.x, x2: P.x, y1: P.y, y2: P.y + P.h, stroke: '#000', 'stroke-width': 1 })
      : svgNode('line', { x1: P.x, x2: P.x + P.w, y1: P.y + P.h, y2: P.y + P.h, stroke: '#000', 'stroke-width': 1 }));
  }
  if (scatter) {
    for (let v = xs.lo; v <= xs.hi + xs.step / 2; v += xs.step) {
      root.append(svgNode('text', { x: xPos(v), y: P.y + P.h + fontPx(axisStyle) + labelGap, 'text-anchor': 'middle',
        ...fontAttrs(axisStyle) }, formatValue(v, null, xs.step)));
    }
  } else if (c.categoryLabels !== false) {
    const angle = c.categoryAngle || 0;
    c.categories.forEach((name, i) => {
      const p = categoryPos(i);
      if (horizontal) {
        root.append(svgNode('text', { x: P.x - labelGap, y: p + fontPx(axisStyle) * 0.35, 'text-anchor': 'end',
          ...fontAttrs(axisStyle) }, name));
      } else if (angle) {
        const y = P.y + P.h + labelGap;
        root.append(svgNode('text', { x: p + fontPx(axisStyle) * 0.35, y, 'text-anchor': 'end',
          transform: `rotate(${-angle} ${p + fontPx(axisStyle) * 0.35} ${y})`, ...fontAttrs(axisStyle) }, name));
      } else {
        root.append(svgNode('text', { x: p, y: P.y + P.h + fontPx(axisStyle) + labelGap, 'text-anchor': 'middle',
          ...fontAttrs(axisStyle) }, name));
      }
    });
  }
}

function drawPie(root, c, P) {
  const s = c.series[0];
  if (!s) return;
  const values = s.values.map((v) => Math.max(0, v || 0));
  const total = values.reduce((a, b) => a + b, 0);
  if (!total) return;
  const cx = P.x + P.w / 2;
  const cy = P.y + P.h / 2;
  const r = Math.min(P.w, P.h) / 2;
  const inner = c.kind === 'doughnut' ? (r * (c.hole || 50)) / 100 : 0;
  const point = (deg, radius) => {
    const a = (deg * Math.PI) / 180; // clockwise from 12 o'clock, as Excel measures it
    return [cx + radius * Math.sin(a), cy - radius * Math.cos(a)];
  };
  let angle = c.rotation || 0;
  values.forEach((v, i) => {
    if (!v) return;
    const sweep = (v / total) * 360;
    const end = angle + sweep;
    const big = sweep > 180 ? 1 : 0;
    const fill = (s.points && s.points[i]) || seriesColour(s, i);
    if (sweep >= 359.99) {
      root.append(svgNode('circle', { cx, cy, r: inner ? (r + inner) / 2 : r, fill: inner ? 'none' : fill,
        stroke: inner ? fill : 'none', 'stroke-width': inner ? r - inner : 0 }));
    } else {
      const [x1, y1] = point(angle, r);
      const [x2, y2] = point(end, r);
      let d = `M${x1},${y1} A${r},${r} 0 ${big} 1 ${x2},${y2}`;
      if (inner) {
        const [x3, y3] = point(end, inner);
        const [x4, y4] = point(angle, inner);
        d += ` L${x3},${y3} A${inner},${inner} 0 ${big} 0 ${x4},${y4} Z`;
      } else {
        d += ` L${cx},${cy} Z`;
      }
      root.append(svgNode('path', { d, fill, stroke: '#ffffff', 'stroke-width': 1 }));
    }
    if (s.labels || s.labelNames) {
      const mid = angle + sweep / 2;
      const radius = c.kind === 'doughnut' ? (r + inner) / 2 : c.outsideLabels ? r * 1.18 : r * 0.62;
      const [x, y] = point(mid, radius);
      const style = s.labelStyle || {};
      const lines = [];
      if (s.labelNames) lines.push(c.categories[i] ?? '');
      if (s.labels === 'percentage') lines.push(`${Math.round((v / total) * 100)}%`);
      else if (s.labels === 'value') lines.push(formatValue(v, c.format));
      const size = fontPx(style, 12);
      const anchor = c.kind === 'pie' && c.outsideLabels ? (Math.sin((mid * Math.PI) / 180) >= 0 ? 'start' : 'end') : 'middle';
      lines.forEach((line, n) => {
        root.append(svgNode('text', { x, y: y + size * 0.35 + (n - (lines.length - 1) / 2) * size * 1.1,
          'text-anchor': anchor, ...fontAttrs(style, 12) }, line));
      });
    }
    angle = end;
  });
}

function drawRadar(root, c, P) {
  const count = Math.max(c.categories.length, ...c.series.map((s) => s.values.length), 3);
  const all = c.series.flatMap((s) => s.values.filter((v) => v !== null));
  const sc = niceScale(Math.min(0, ...all), Math.max(...all));
  const cx = P.x + P.w / 2;
  const cy = P.y + P.h / 2;
  const r = Math.min(P.w, P.h) / 2;
  const at = (i, v) => {
    const a = (i / count) * 2 * Math.PI;
    const d = ((v - sc.lo) / (sc.hi - sc.lo)) * r;
    return [cx + d * Math.sin(a), cy - d * Math.cos(a)];
  };
  for (let v = sc.lo + sc.step; v <= sc.hi + sc.step / 2; v += sc.step) {
    const ring = Array.from({ length: count }, (_, i) => at(i, v).join(',')).join(' ');
    root.append(svgNode('polygon', { points: ring, fill: 'none', stroke: c.gridlines || '#d0d0d0', 'stroke-width': 0.75 }));
  }
  c.series.forEach((s) => {
    const points = s.values.map((v, i) => at(i, v || 0).join(',')).join(' ');
    root.append(svgNode('polygon', { points, fill: s.color || 'none', 'fill-opacity': 0.5,
      stroke: s.line || s.color || '#4472C4', 'stroke-width': 1.5 }));
  });
  c.categories.forEach((name, i) => {
    const [x, y] = at(i, sc.hi + sc.step * 0.25);
    root.append(svgNode('text', { x, y, 'text-anchor': 'middle', ...fontAttrs(c.axisStyle) }, name));
  });
}

function drawLegend(root, c, L) {
  const byPoint = c.series.length === 1 && c.series[0].points;
  const entries = byPoint
    ? c.categories.map((name, i) => [name, c.series[0].points[i] || '#4472C4'])
    : c.series.map((s, i) => [s.name, seriesColour(s, i)]);
  const style = c.legendStyle;
  const size = fontPx(style);
  const swatch = size * 0.65;
  const lineHeight = size * 1.45;
  const g = svgNode('g');
  root.append(g);
  if (L.w >= L.h * 2) {
    // A wide legend lists its entries across, in as few rows as fit, each column as wide as its
    // widest entry.
    const width = (name) => textWidth(name, style) + swatch + 14;
    let columns = [];
    for (let perRow = entries.length; perRow >= 1; perRow--) {
      columns = Array.from({ length: perRow }, (_, j) =>
        Math.max(...entries.filter((_, i) => i % perRow === j).map(([name]) => width(name))));
      if (columns.reduce((a, b) => a + b, 0) <= L.w) break;
    }
    const perRow = columns.length;
    const starts = columns.map((_, j) => columns.slice(0, j).reduce((a, b) => a + b, 0));
    const rows = Math.ceil(entries.length / perRow);
    const left = L.x + Math.max(0, (L.w - columns.reduce((a, b) => a + b, 0)) / 2);
    const top = L.y + (L.h - rows * lineHeight) / 2;
    entries.forEach(([name, colour], i) => {
      const x = left + starts[i % perRow];
      const y = top + Math.floor(i / perRow) * lineHeight + lineHeight / 2;
      g.append(svgNode('rect', { x, y: y - swatch / 2, width: swatch, height: swatch, fill: colour }));
      g.append(svgNode('text', { x: x + swatch + 4, y: y + size * 0.35, ...fontAttrs(style) }, name));
    });
  } else {
    const top = L.y + Math.max(0, (L.h - entries.length * lineHeight) / 2);
    entries.forEach(([name, colour], i) => {
      const y = top + i * lineHeight + lineHeight / 2;
      g.append(svgNode('rect', { x: L.x, y: y - swatch / 2, width: swatch, height: swatch, fill: colour }));
      g.append(svgNode('text', { x: L.x + swatch + 4, y: y + size * 0.35, ...fontAttrs(style) }, name));
    });
  }
}
