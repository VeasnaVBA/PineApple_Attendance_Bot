const { webhookCallback } = require('grammy');
const bot = require('../src/bot');

module.exports = webhookCallback(bot, 'http');
