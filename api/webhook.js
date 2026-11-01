const bot = require('../src/bot');

module.exports = async (req, res) => {
  if (req.method === 'POST') {
    try {
      if (!bot.isInited()) {
        await bot.init();
      }
      const body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
      if (body) {
        await bot.handleUpdate(body);
      }
      return res.status(200).send('OK');
    } catch (err) {
      console.error('Error handling webhook update:', err);
      // Always return 200 to Telegram so it doesn't repeatedly retry failed updates!
      return res.status(200).send('OK');
    }
  }

  // GET request (Health check / ping)
  return res.status(200).json({
    status: 'active',
    message: '🍍 PineApple Attendance Telegram Webhook is ready!'
  });
};
