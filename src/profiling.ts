/** BEGIN src/profiling.ts */
namespace Profiler {
  const _executionTimes: {
    [key: string]: {
      executionTime: number | null;
      calls: number;
      cumulativeExecutionTime: number;
    };
  } = {};

  export function wrap(func: Function, name: string = func.name) {
    if (!_executionTimes[name]) {
      _executionTimes[name] = {
        executionTime: null,
        calls: 0,
        cumulativeExecutionTime: 0,
      };
    }

    return (...args: any[]) => {
      const start = Date.now();
      const output = func(...args);
      const end = Date.now();
      const executionTime = end - start;
      _executionTimes[name].executionTime = executionTime;
      _executionTimes[name].calls += 1;
      _executionTimes[name].cumulativeExecutionTime += executionTime;
      return output;
    };
  }

  export function format(): string {
    const keys = Object.keys(_executionTimes);
    const sortedKeys = keys.sort((a, b) => {
      const aTime = _executionTimes[a].cumulativeExecutionTime;
      const bTime = _executionTimes[b].cumulativeExecutionTime;
      return bTime - aTime;
    });

    let output = "";
    for (const key of sortedKeys) {
      const { executionTime, calls, cumulativeExecutionTime } =
        _executionTimes[key];
      output += `${key}: ${executionTime}ms (cumulative: ${cumulativeExecutionTime}ms, calls: ${calls})\n`;
    }

    return output;
  }
}
/** END src/profiling.ts */
