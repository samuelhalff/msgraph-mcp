import winston from 'winston';

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

export default logger;
