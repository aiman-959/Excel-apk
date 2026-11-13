import { useId } from 'react';

/** Petak mark (assets/brand/petak-logo.svg) as an inline component: 9 "plots" on a dark green tile. */
const PLOTS: Array<[number, number, string]> = [
  [44, 44, '#16A34A'], [98, 44, '#22C55E'], [152, 44, '#15803D'],
  [44, 98, '#22C55E'], [98, 98, '#15803D'], [152, 98, '#FBBF24'],
  [44, 152, '#15803D'], [98, 152, '#16A34A'], [152, 152, '#22C55E'],
];

export function PetakLogo({ size = 32, tile = true, title }: { size?: number; tile?: boolean; title?: string }) {
  const id = useId().replace(/:/g, '');
  return (
    <svg width={size} height={size} viewBox={tile ? '0 0 240 240' : '40 40 160 160'} role={title ? 'img' : undefined} aria-hidden={title ? undefined : true} aria-label={title}>
      <defs>
        <pattern id={`rows${id}`} width="8" height="8" patternUnits="userSpaceOnUse">
          <line x1="0" y1="4" x2="8" y2="4" stroke="#FFFFFF" strokeOpacity="0.28" strokeWidth="2" />
        </pattern>
      </defs>
      {tile && <rect width="240" height="240" rx="56" fill="#14532D" />}
      {PLOTS.map(([x, y, c]) => (
        <rect key={`${x}-${y}`} x={x} y={y} width="44" height="44" rx="8" fill={c} />
      ))}
      <g fill={`url(#rows${id})`}>
        {PLOTS.map(([x, y]) => (
          <rect key={`p${x}-${y}`} x={x} y={y} width="44" height="44" rx="8" />
        ))}
      </g>
    </svg>
  );
}
