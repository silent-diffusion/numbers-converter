'use strict';

// Draws a text box, shape or picture for the preview from the helper's description of it (see
// describe() in py/drawings.py), as Excel will show it: placed and turned the same way, with the
// same outline, fill and line, and its text in the same fonts and alignment. Built with DOM calls
// only (svgNode is in charts.js), so text from the file can never become markup.

const DRAWING_PT = 4 / 3; // fonts, line widths and spacing are in points
const DASHES = { sysDot: [1, 1], dash: [4, 3], lgDash: [8, 3] }; // in line widths
let gradientCount = 0;

function drawDrawing(d) {
  const box = document.createElement('div');
  box.className = `drawing ${d.kind}`;
  Object.assign(box.style, { left: `${d.x}px`, top: `${d.y}px`, width: `${d.w}px`, height: `${d.h}px` });
  if (d.rotation) box.style.transform = `rotate(${d.rotation}deg)`;
  if (d.kind === 'image') {
    box.append(picture(d));
    if (d.line) box.append(outline(d, false));
  } else {
    box.append(outline(d, true));
    if (d.paragraphs) box.append(textBlock(d));
  }
  return box;
}

// The outline's path in pixels, stretched from its own box to the drawing's.
function scaledPath(d) {
  const sx = d.w / d.box[0];
  const sy = d.h / d.box[1];
  return d.path.map(([cmd, ...values]) => cmd + values.map((v, i) => (i % 2 ? v * sy : v * sx).toFixed(2)).join(' ')).join(' ');
}

function outline(d, filled) {
  const svg = svgNode('svg', { class: 'outline', width: Math.max(d.w, 1), height: Math.max(d.h, 1) });
  const flips = [d.flipH && `translate(${d.w} 0) scale(-1 1)`, d.flipV && `translate(0 ${d.h}) scale(1 -1)`].filter(Boolean);
  const group = svgNode('g', { transform: flips.join(' ') || null });
  let fill = 'none';
  if (filled && d.fill?.color) fill = d.fill.color;
  else if (filled && d.fill?.stops) {
    const id = `drawing-gradient-${++gradientCount}`;
    const gradient = d.fill.radial
      ? svgNode('radialGradient', { id })
      : svgNode('linearGradient', { id, x1: 0, y1: 0.5, x2: 1, y2: 0.5, gradientTransform: `rotate(${d.fill.angle} 0.5 0.5)` });
    for (const [offset, colour] of d.fill.stops) gradient.append(svgNode('stop', { offset, 'stop-color': colour }));
    const defs = svgNode('defs');
    defs.append(gradient);
    svg.append(defs);
    fill = `url(#${id})`;
  }
  const line = d.line ? {
    stroke: d.line.color,
    'stroke-width': Math.max(0.75, d.line.width),
    'stroke-dasharray': DASHES[d.line.dash] ? DASHES[d.line.dash].map((n) => n * d.line.width).join(' ') : null,
  } : { stroke: 'none' };
  let shape;
  if (d.outline === 'line') shape = svgNode('line', { x1: 0, y1: 0, x2: d.w, y2: d.h, ...line });
  else if (d.outline === 'path' && d.path) shape = svgNode('path', { d: scaledPath(d), fill, ...line });
  else shape = svgNode('rect', { x: 0, y: 0, width: d.w, height: d.h, fill, ...line });
  group.append(shape);
  svg.append(group);
  return svg;
}

function textBlock(d) {
  const block = document.createElement('div');
  block.className = `text ${d.anchor}`;
  const [left, top, right, bottom] = d.padding;
  block.style.padding = `${top}px ${right}px ${bottom}px ${left}px`;
  for (const p of d.paragraphs) {
    const para = document.createElement('p');
    para.style.textAlign = { l: 'left', ctr: 'center', r: 'right', just: 'justify' }[p.align] || 'left';
    para.style.marginTop = `${p.before * DRAWING_PT}px`;
    para.style.marginBottom = `${p.after * DRAWING_PT}px`;
    para.style.paddingLeft = `${p.left * DRAWING_PT}px`;
    if (p.first && !p.bullet) para.style.textIndent = `${p.first * DRAWING_PT}px`;
    para.style.lineHeight = p.lineHeight ? `${p.lineHeight * DRAWING_PT}px` : String(1.2 * (p.lineScale || 1));
    para.style.fontSize = `${p.size * DRAWING_PT}px`;
    if (p.bullet) {
      const mark = document.createElement('span');
      mark.className = 'bullet';
      mark.textContent = p.bullet;
      para.append(mark);
    }
    for (const [text, style] of p.runs) {
      text.split('\u2028').forEach((piece, i) => {
        if (i) para.append(document.createElement('br'));
        if (piece) para.append(run(piece, style));
      });
    }
    if (!p.runs.length) para.append('\u00a0'); // an empty paragraph still takes a line
    block.append(para);
  }
  return block;
}

function run(text, s) {
  const span = document.createElement('span');
  span.textContent = text;
  const st = span.style;
  st.fontFamily = `"${s.font || 'Arial'}", Arial, sans-serif`;
  st.fontSize = `${(s.size || 12) * DRAWING_PT * (s.script ? 0.7 : 1)}px`;
  if (s.b) st.fontWeight = '700';
  if (s.i) st.fontStyle = 'italic';
  const decoration = [s.u && 'underline', s.st && 'line-through'].filter(Boolean).join(' ');
  if (decoration) st.textDecoration = decoration;
  if (s.color) st.color = s.alpha !== undefined && s.alpha < 1 ? hexToRgba(s.color, s.alpha) : s.color;
  if (s.script) st.verticalAlign = s.script === 'super' ? 'super' : 'sub';
  if (s.caps === 'all') st.textTransform = 'uppercase';
  if (s.caps === 'small') st.fontVariant = 'small-caps';
  return span;
}

function hexToRgba(hex, alpha) {
  const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16));
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

// A picture, cropped and clipped to its outline as in Excel; just its name where the preview
// doesn't carry the picture (it's still in the saved file).
function picture(d) {
  const frame = document.createElement('div');
  frame.className = 'crop';
  if (d.outline === 'path' && d.path) frame.style.clipPath = `path("${scaledPath(d)}")`;
  if (!d.src) {
    frame.classList.add('missing');
    frame.textContent = d.name || 'Picture';
    return frame;
  }
  const img = document.createElement('img');
  img.src = d.src;
  img.alt = d.name || '';
  img.draggable = false;
  const [l, t, r, b] = d.crop || [0, 0, 0, 0];
  const width = d.w / Math.max(0.01, 1 - l - r);
  const height = d.h / Math.max(0.01, 1 - t - b);
  Object.assign(img.style, { width: `${width}px`, height: `${height}px`, left: `${-l * width}px`, top: `${-t * height}px` });
  const flips = [d.flipH && 'scaleX(-1)', d.flipV && 'scaleY(-1)'].filter(Boolean);
  if (flips.length) frame.style.transform = flips.join(' ');
  frame.append(img);
  return frame;
}
