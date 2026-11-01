module.exports = (req, res) => {
  res.status(200).json({
    status: 'online',
    message: '🍍 PineApple Attendance Bot is running on Vercel!',
    webhook: '/api/webhook'
  });
};
