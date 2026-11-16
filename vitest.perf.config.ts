import { defineConfig } from 'vitest/config';

// The time limits of the engine (R46, R47), which `make perf` checks apart
// from `make check`: a shared CI machine is not the laptop they are set for.
export default defineConfig({
  test: { environment: 'node', include: ['tests/perf/*.test.ts'] },
});
