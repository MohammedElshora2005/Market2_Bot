const { Telegraf, Markup } = require('telegraf');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
require('dotenv').config();

const bot = new Telegraf(process.env.BOT_TOKEN);

const serviceAccountAuth = new JWT({
  email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
  key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);

// قاعدة بيانات وهمية للمنتجات لتطوير البوت
const PRODUCTS = [
  { id: 'p1', name: 'كوكاكولا عائلي', price: 40, desc: 'زجاجة كوكاكولا 2.25 لتر منعشة' },
  { id: 'p2', name: 'شيبسي عائلي طماطم', price: 25, desc: 'كيس شيبسي حجم عائلي طعم الطماطم' },
  { id: 'p3', name: 'جبنة دومتي فيتا', price: 55, desc: 'علبة جبنة دومتي فيتا 500 جرام' }
];

async function saveOrderToSheet(orderData) {
  try {
    await doc.loadInfo();
    const sheet = doc.sheetsByTitle['Orders'];
    await sheet.addRow({
      'ID': orderData.chatId,
      'اسم العميل': orderData.customerName,
      'الطلبات': orderData.items,
      'الإجمالي': orderData.total,
      'العنوان': orderData.address,
      'رقم الهاتف': orderData.phone,
      'التاريخ': orderData.date,
      'عدد العناصر': orderData.itemCount,
      'رقم الطلب': orderData.orderNumber,
      'الحالة': 'جاري التنفيذ'
    });
    return true;
  } catch (error) {
    console.error(error);
    return false;
  }
}

bot.start((ctx) => {
  const firstName = ctx.from.first_name || 'يا فنان';
  ctx.replyWithHTML(
    `أهلاً بك يا <b>${firstName}</b> في سوبر ماركت بوب مارت! 👋\n\nاضغط على الزر بالأسفل لتصفح المنتجات المتوفرة وطلبها فوراً.`,
    Markup.inlineKeyboard([
      [Markup.button.callback('🛒 تصفح المنتجات', 'view_products')]
    ])
  );
});

bot.action('view_products', async (ctx) => {
  await ctx.answerCbQuery();
  await ctx.reply('📦 إليك المنتجات المتاحة لدينا اليوم:');
  
  PRODUCTS.forEach(async (product) => {
    const message = `<b>📦 ${product.name}</b>\n📝 ${product.desc}\n💰 السعر: <code>${product.price} EGP</code>`;
    await ctx.replyWithHTML(
      message,
      Markup.inlineKeyboard([
        [Markup.button.callback(`🛒 اطلب ${product.name}`, `buy_${product.id}`)]
      ])
    );
  });
});

PRODUCTS.forEach((product) => {
  bot.action(`buy_${product.id}`, async (ctx) => {
    await ctx.answerCbQuery();
    const orderNumber = Math.floor(100000 + Math.random() * 900000);
    
    const orderData = {
      chatId: ctx.from.id,
      customerName: ctx.from.first_name || 'عميل بوت',
      items: product.name,
      total: `${product.price} EGP`,
      address: 'تحديد تلقائي عبر البوت',
      phone: '01000000000',
      date: new Date().toLocaleString('ar-EG'),
      itemCount: 1,
      orderNumber: orderNumber
    };

    await ctx.reply(`⏳ جاري تسجيل طلبك لـ (${product.name}) في الجداول...`);
    const success = await saveOrderToSheet(orderData);

    if (success) {
      ctx.replyWithHTML(`✅ <b>تم تسجيل طلبك بنجاح!</b>\n📦 رقم الطلب: <code>#${orderNumber}</code>\n\nتابع الشيت الخاص بك لمشاهدة الطلب الجديد.`);
    } else {
      ctx.reply('❌ حدث خطأ أثناء تسجيل الطلب، يرجى المحاولة مرة أخرى.');
    }
  });
});

module.exports = async (req, res) => {
  try {
    if (req.method === 'POST') {
      await bot.handleUpdate(req.body);
      res.status(200).send('OK');
    } else {
      res.status(200).send('سيرفر البوت يعمل بنجاح!');
    }
  } catch (error) {
    console.error(error);
    res.status(500).send('Internal Server Error');
  }
};
