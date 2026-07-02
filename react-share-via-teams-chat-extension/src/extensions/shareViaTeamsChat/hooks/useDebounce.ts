import * as React from 'react';

/**
 * Debounces any value by `delay` milliseconds. The returned value
 * only updates after the input has remained stable for `delay` ms.
 */
export function useDebounce<T>(value: T, delay: number): T {
  const [debounced, setDebounced] = React.useState<T>(value);

  React.useEffect(() => {
    const handle = setTimeout(() => setDebounced(value), delay);
    return () => clearTimeout(handle);
  }, [value, delay]);

  return debounced;
}
