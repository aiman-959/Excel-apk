import { describe, expect, it } from 'vitest';
import { droppedFraction, nextFrameLevel, thermalLevel } from '../../src/device/governor';

describe('load governor', () => {
  it('counts frames that missed their refresh slot', () => {
    expect(droppedFraction(Array(10).fill(8.3), 120)).toBe(0);
    expect(droppedFraction([8.3, 8.3, 16.7, 25, 8.3, 8.3, 8.3, 8.3], 120)).toBe(0.25);
    expect(droppedFraction(Array(10).fill(16.7), 60)).toBe(0);
  });
  it('goes busy quickly and recovers only after a calm period', () => {
    expect(nextFrameLevel(0, 0.4, 0)).toBe(1);
    expect(nextFrameLevel(1, 0.05, 2000)).toBe(1);
    expect(nextFrameLevel(1, 0.05, 7000)).toBe(0);
    expect(nextFrameLevel(0, 0.2, 0)).toBe(0);
  });
  it('maps Android thermal status', () => {
    expect([0, 1, 2, 3, 4, 6].map(thermalLevel)).toEqual([0, 0, 1, 2, 2, 2]);
  });
});
