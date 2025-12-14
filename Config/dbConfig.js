module.exports = {
  server: 'localhost\\SQLEXPRESS',   // MUST be named instance
  database: 'OP',
  driver: 'msnodesqlv8',

  options: {
    trustedConnection: true,         // <-- KEY FIX
    enableArithAbort: true
  }
};
