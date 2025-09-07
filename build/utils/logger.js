export const logger = (namespace) => {
    const log = (level, ...args) => {
        const timestamp = new Date().toISOString();
        const prefix = `[${timestamp}] [${namespace}] [${level.toUpperCase()}]`;
        console.log(prefix, ...args);
    };
    return {
        info: (...args) => log('info', ...args),
        warn: (...args) => log('warn', ...args),
        error: (...args) => log('error', ...args),
        debug: (...args) => log('debug', ...args)
    };
};
