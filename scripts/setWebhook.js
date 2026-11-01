require('dotenv').config();
const { Bot } = require('grammy');

const token = process.env.BOT_TOKEN;
const vercelUrl = process.argv[2] || process.env.VERCEL_URL;

if (!token) {
  console.error('❌ BOT_TOKEN is missing in .env!');
  process.exit(1);
}

if (!vercelUrl) {
  console.log('📌 Usage: node scripts/setWebhook.js <YOUR_VERCEL_URL>');
  console.log('Example: node scripts/setWebhook.js https://pineapple-attendance-bot.vercel.app');
  process.exit(1);
}

const cleanUrl = vercelUrl.startsWith('http') ? vercelUrl : `https://${vercelUrl}`;
const webhookUrl = `${cleanUrl.replace(/\/$/, '')}/api/webhook`;

const bot = new Bot(token);

(async () => {
  console.log(`🔗 Setting Telegram Webhook to: ${webhookUrl}...`);
  try {
    await bot.api.setWebhook(webhookUrl);
    console.log('✅ Webhook set successfully!');
    const info = await bot.api.getWebhookInfo();
    console.log('Webhook Info:', info);
  } catch (err) {
    console.error('❌ Failed to set webhook:', err.message);
  }
})();
