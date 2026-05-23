const { Telegraf } = require('telegraf');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
require('dotenv').config();

// 1. تهيئة بوت تليجرام باستخدام التوكن من متغيرات البيئة
const bot = new Telegraf(process.env.BOT_TOKEN);

// 2. إعداد الاتصال بجوجل شيت باستخدام الحساب الخدمي (Service Account)
const serviceAccountAuth = new JWT({
  email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
  key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);

// دالة مساعدة لحفظ الطلب في جوجل شيت
async function saveOrderToSheet(orderData) {
  try {
    await doc.loadInfo();
    const sheet = doc.sheetsByTitle['Orders'];
    
    // إضافة صف جديد بالترتيب المتوافق تماماً مع أعمدة الشيت الخاص بك
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
      'الحالة': 'جاري التنفيذ' // الحالة الافتراضية لأي طلب جديد
    });
    console.log(`✅ تم تسجيل الطلب رقم ${orderData.orderNumber} بنجاح في الشيت.`);
    return true;
  } catch (error) {
    console.error('❌ خطأ أثناء التسجيل في جوجل شيت:', error);
    return false;
  }
}

// 3. معالجة أوامر البوت والتفاعل مع المستخدم
bot.start((ctx) => {
  const firstName = ctx.from.first_name || 'يا فنان';
  ctx.replyWithHTML(`أهلاً بك يا <b>${firstName}</b> في سوبر ماركت بوب مارت! 👋\n\nتم ربط حسابك بنجاح. معرف الحساب الخاص بك (ID) هو: <code>${ctx.from.id}</code>\n\nيمكنك الآن إرسال طلباتك من خلال التطبيق أو الموقع وسيصلك إشعار فوري هنا عند تحديث حالة طلبك! ✨`);
});

// مثال بسيط لاستقبال طلب تجريبي وضخه في الشيت مباشرة (للتأكد من سلامة الربط)
bot.command('testorder', async (ctx) => {
  const orderNumber = Math.floor(100000 + Math.random() * 900000); // توليد رقم طلب عشوائي من 6 أرقام
  const orderData = {
    chatId: ctx.from.id,
    customerName: ctx.from.first_name || 'عميل تجريبي',
    items: 'كوكاكولا، شيبسي عائلي، جبنة دومتي',
    total: '120 EGP',
    address: 'أشمون، المنوفية',
    phone: '01000000000',
    date: new Date().toLocaleString('ar-EG'),
    itemCount: 3,
    orderNumber: orderNumber
  };

  ctx.reply('⏳ جاري تسجيل طلب تجريبي في الجداول للتأكد من الربط...');
  const success = await saveOrderToSheet(orderData);
  
  if (success) {
    ctx.replyWithHTML(`✅ <b>تم تسجيل الطلب التجريبي بنجاح!</b>\n📦 رقم الطلب: <code>#${orderNumber}</code>\n\nراجع الجوجل شيت الآن ستجده ظهر في صفحة Orders تلقائياً.`);
  } else {
    ctx.reply('❌ فشل تسجيل الطلب. تأكد من إعداد متغيرات البيئة (Environment Variables) وصلاحيات الحساب الخدمي.');
  }
});

// 4. إعداد مسار الـ Webhook الخاص بـ Vercel ليعمل كدالة Serverless
module.exports = async (req, res) => {
  try {
    if (req.method === 'POST') {
      // تمرير التحديث القادم من تليجرام إلى مكتبة Telegraf لتعالجه
      await bot.handleUpdate(req.body);
      res.status(200).send('OK');
    } else {
      res.status(200).send('سيرفر البوت يعمل بنجاح! بانتظار تحديثات تليجرام عبر POST.');
    }
  } catch (error) {
    console.error('خطأ في معالجة طلب الـ Webhook:', error);
    res.status(500).send('Internal Server Error');
  }
};
