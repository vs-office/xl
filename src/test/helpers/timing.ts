// Wall-clock limits are tuned for a developer machine; shared CI runners are slower and noisier.
const SCALE = process.env.CI ? 5 : 1;

export const budget = (ms: number): number => ms * SCALE;
