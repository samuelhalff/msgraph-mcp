export const logger = (namespace: string) => {
  const log = (level: string, ...args: unknown[]) => {
    const timestamp = new Date().toISOString();
    const prefix = `[${timestamp}] [${namespace}] [${level.toUpperCase()}]`;
    console.log(prefix, ...args);
  };

  return {
    info: (...args: unknown[]) => log('info', ...args),
    warn: (...args: unknown[]) => log('warn', ...args),
    error: (...args: unknown[]) => log('error', ...args),
    debug: (...args: unknown[]) => log('debug', ...args)
  };
};