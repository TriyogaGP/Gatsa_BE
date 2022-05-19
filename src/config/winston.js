const appRoot = require('app-root-path');
const winston = require('winston');
const { SqlTransport } = require('winston-sql-transport');
const moment = require('moment');
const dotenv = require('dotenv');
dotenv.config();

const transportConfig = {
  client: 'mysql2',
  connection: {
    host: process.env.DB_HOST || "localhost",
    user: process.env.DB_USER || "root",
    password: process.env.DB_PASS || "",
    database: process.env.DB_DATABASE || "db_yoga",
    port: process.env.DB_PORT || "3307"
  },
  tableName: 'history_logs',
};

// Gunanya buat nyetting log yang akan dikeluarin, baik itu ke file berupa output maupun console terminal 
const today = moment();
const options = {
  file: {
    level: 'info',
    filename: `${appRoot.path}/src/logs/logger_history.log`,
    handleExceptions: true,
    json: true,
    maxsize: 5242880, //ukuran file maksimal 5MB
    maxFiles: 5,
    colorize: false,
    format: winston.format.combine(
      winston.format.timestamp({
        format: 'YYYY-MM-DD HH:mm:ss'
      }),
      winston.format.json(),
    ) 
  },
  console: {
    level: 'debug',
    handleExceptions: true,
    json: false,
    colorize: true,
    format: winston.format.combine(
      winston.format.timestamp({
        format: 'YYYY-MM-DD HH:mm:ss'
      }),
      winston.format.json(),
    )
  },
};

// Panggil class si winston dengan setting yang udah kita buat
const logger = winston.createLogger({
  format: winston.format.combine(
    winston.format.timestamp({
      format: 'YYYY-MM-DD HH:mm:ss'
    }),
    // winston.format.metadata({ fill: ['message', 'level', 'timestamp'] }),
    winston.format.json(),
  ),
  transports: [
    new winston.transports.File(options.file),
    new winston.transports.Console(options.console),
    new SqlTransport(transportConfig)
  ],
  exitOnError: false, // Aplikasi gabakalan berhenti kalo ada exception
});

// Bikin file stream (nulis file) yang dimana bakalan dipake sama morgan (sm*ash) ups hahaha.`
logger.stream = {
  write: function(message, encoding) {
    // pake log level info aja supaya outputnya dipake sama file stream dan console.
    logger.info(message);
  },
};

// const sendLogDB = winston.createLogger({
//   format: winston.format.combine(
//     winston.format.timestamp({
//       format: 'YYYY-MM-DD HH:mm:ss'
//     }),
//     // winston.format.metadata({ fill: ['message', 'level', 'timestamp'] }),
//     winston.format.json(),
//   ),
//   transports: [new SqlTransport(transportConfig)],
// });

// sendLogDB.stream = {
//   write: function(message, encoding) {
//     // pake log level info aja supaya outputnya dipake sama file stream dan console.
//     sendLogDB.info(message);
//   },
// };

// (async () => {
//   const transport = new SqlTransport(transportConfig);
//   await transport.init();
// })();

module.exports = {
  logger,
  // sendLogDB
};