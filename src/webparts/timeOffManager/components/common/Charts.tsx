import * as React from 'react';

/** Single-series magnitude colour (same as the Employee Directory dashboards). */
export const SERIES_COLOR = '#2a78d6';

export interface IDatum {
  label: string;
  value: number;
  color?: string;
}

export function BarList(props: { items: IDatum[]; format?: (v: number) => string; empty?: string }): JSX.Element {
  const max = Math.max(1, ...props.items.map((i) => i.value));
  if (!props.items.length) return <div className="to-muted to-small">{props.empty || 'No data yet.'}</div>;
  return (
    <div className="to-bars">
      {props.items.map((d, i) => (
        <div key={d.label + i} className="to-bar-row">
          <span className="to-bar-label" title={d.label}>
            {d.label}
          </span>
          <span className="to-bar-track">
            <span className="to-bar-fill" style={{ display: 'block', width: `${(d.value / max) * 100}%`, background: d.color || SERIES_COLOR }} />
          </span>
          <span className="to-bar-value">{props.format ? props.format(d.value) : d.value}</span>
        </div>
      ))}
    </div>
  );
}

/** Vertical columns over evenly spaced categories (e.g. months), values labelled. */
export function Columns(props: { labels: string[]; values: number[]; height?: number; color?: string }): JSX.Element {
  const w = 640;
  const h = props.height || 190;
  const pad = { l: 8, r: 8, t: 18, b: 22 };
  const n = props.labels.length;
  const max = Math.max(1, ...props.values);
  const slot = (w - pad.l - pad.r) / Math.max(1, n);
  const bw = Math.min(34, slot * 0.62);
  return (
    <div className="to-chart">
      <svg viewBox={`0 0 ${w} ${h}`} role="img" aria-label={props.labels.map((l, i) => `${l}: ${props.values[i]}`).join(', ')}>
        <line x1={pad.l} x2={w - pad.r} y1={h - pad.b} y2={h - pad.b} stroke="#dfe5e6" />
        {props.values.map((v, i) => {
          const bh = ((h - pad.t - pad.b) * v) / max;
          const x = pad.l + slot * i + (slot - bw) / 2;
          return (
            <g key={i}>
              <rect x={x} y={h - pad.b - bh} width={bw} height={Math.max(0, bh)} rx={4} fill={props.color || SERIES_COLOR} />
              {v > 0 && (
                <text x={x + bw / 2} y={h - pad.b - bh - 5} textAnchor="middle" style={{ fill: '#17232e', fontWeight: 600 }}>
                  {Math.round(v * 10) / 10}
                </text>
              )}
              <text x={x + bw / 2} y={h - 6} textAnchor="middle">
                {props.labels[i]}
              </text>
            </g>
          );
        })}
      </svg>
    </div>
  );
}
