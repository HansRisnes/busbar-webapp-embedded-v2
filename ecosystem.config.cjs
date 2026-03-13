module.exports = {
  apps: [
    {
      name: "busbar-api",
      script: "server/mail.js",
      cwd: __dirname,
      instances: 1,
      autorestart: true,
      max_restarts: 20,
      watch: false,
      env: {
        NODE_ENV: "production",
        PORT: 5500
      }
    }
  ]
};
