var path = require('path'),
    rootPath = path.normalize(__dirname + '/..'),
    env = process.env.NODE_ENV || 'development';

var config = {
  development: {
    root: rootPath,
    app: {
      name: 'o365-nodejs-customaction'
    },
    port: process.env.PORT || 8443,
  },

  test: {
    root: rootPath,
    app: {
      name: 'o365-nodejs-customaction'
    },
    port: process.env.PORT || 8443,
  },

  production: {
    root: rootPath,
    app: {
      name: 'o365-nodejs-customaction'
    },
    port: process.env.PORT || 8443,
  }
};

module.exports = config[env];
