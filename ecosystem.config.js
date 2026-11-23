module.exports = {
  apps: [
    {
      name: 'sales-dashboard',
      script: 'server.js',
      cwd: './next-dashboard',
      max_memory_restart: '500M',
      env: {
        NODE_ENV: 'production',
        PORT: 3000
      }
    }
  ]
};
