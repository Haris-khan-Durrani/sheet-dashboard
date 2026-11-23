/**
 * Root Server Entrypoint
 * Allows running `node server.js` or `pm2 start server.js` directly from the repository root.
 */
const path = require('path');
const targetDir = path.join(__dirname, 'next-dashboard');

try {
  process.chdir(targetDir);
} catch (e) {}

require(path.join(targetDir, 'server.js'));
