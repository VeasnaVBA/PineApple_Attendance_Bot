require('dotenv').config();
const { Bot } = require('grammy');

const token = process.env.BOT_TOKEN;
if (!token) {
  console.error('❌ BOT_TOKEN is missing in .env!');
  process.exit(1);
}

const bot = new Bot(token);

(async () => {
  console.log('🔄 Deleting Telegram Webhook to enable local polling...');
  try {
    await bot.api.deleteWebhook();
    console.log('✅ Webhook deleted! Bot can now run locally using polling (node src/bot.js).');
  } catch (err) {
    console.error('❌ Failed to delete webhook:', err.message);
  }
})();
