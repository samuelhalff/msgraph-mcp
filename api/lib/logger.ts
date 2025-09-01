import winston from 'winston';
import jwt from 'jsonwebtoken';

// Create a logger instance
const logger = winston.createLogger({
  level: process.env.LOG_LEVEL || 'info',
  format: winston.format.combine(
    winston.format.timestamp(),
    winston.format.errors({ stack: true }),
    winston.format.json()
  ),
  defaultMeta: { service: 'msgraph-mcp' },
  transports: [],
  exceptionHandlers: [],
  rejectionHandlers: [],
});

// Add file transports with error handling
try {
  // Write all logs with importance level of `error` or less to `error.log`
  logger.add(new winston.transports.File({
    filename: 'logs/error.log',
    level: 'error',
    handleExceptions: true,
    handleRejections: true
  }));

  // Write all logs with importance level of `info` or less to `combined.log`
  logger.add(new winston.transports.File({
    filename: 'logs/combined.log',
    handleExceptions: true,
    handleRejections: true
  }));
} catch (error) {
  console.warn('Failed to initialize file logging, falling back to console only:', error instanceof Error ? error.message : String(error));
}

// Always add console transport
logger.add(new winston.transports.Console({
  format: winston.format.combine(
    winston.format.colorize(),
    winston.format.simple()
  ),
  handleExceptions: true,
  handleRejections: true
}));

// JWT Token logging helper function
export function logToken(token: string, context: string) {
  try {
    const decoded = jwt.decode(token, { complete: true });
    if (decoded && decoded.payload) {
      const payload = decoded.payload as any;
      logger.info(`[${context}] Token received`, {
        aud: payload.aud,
        scp: payload.scp,
        exp: payload.exp ? new Date(payload.exp * 1000).toISOString() : 'N/A',
        iat: payload.iat ? new Date(payload.iat * 1000).toISOString() : 'N/A',
        iss: payload.iss,
        sub: payload.sub,
        tokenType: payload.token_type || 'N/A'
      });
    } else {
      logger.warn(`[${context}] Token could not be decoded - invalid format`);
    }
  } catch (err) {
    logger.error(`[${context}] Failed to decode token`, { error: err instanceof Error ? err.message : String(err) });
  }
}

export default logger;
