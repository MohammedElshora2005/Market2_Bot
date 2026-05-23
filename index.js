import { Telegraf } from 'telegraf';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';
import axios from 'axios';
import 'dotenv/config';

const bot = new Telegraf(process.env.TELEGRAM_BOT_TOKEN);

// تخزين جلسات المستخدمين (في الذاكرة)
const userSessions = {};

async function getDoc() {
    const rawKey = process.env.GOOGLE_PRIVATE_KEY || '';
    const formattedKey = rawKey.replace(/\\n/g, '\n').replace(/"/g, '');

    const serviceAccountAuth = new JWT({
        email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        key: formattedKey,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(process.env.SHEET_ID, serviceAccountAuth);
    await doc.loadInfo();
    return doc;
}

// أمر /start
bot.start(async (ctx) => {
    const userId = ctx.from.id;
    userSessions[userId] = { step: 'idle' };
    await ctx.reply('✅ البوت شغال يا فندم!\n\n📋 الأوامر:\n/products - عرض المنتجات\n/cancel - إلغاء الطلب');
});

// أمر عرض المنتجات
bot.command('products', async (ctx) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['Products'];
        await sheet.loadHeaderRow();
        const rows = await sheet.getRows();
        
        let message = '🛒 منتجاتنا:\n\n';
        for (const row of rows) {
            const product = row.get('المنتج');
            const price = row.get('السعر');
            const unit = row.get('الوحدة');
            if (product && price) {
                message += `• ${product}: ${price} جنيه / ${unit || 'وحدة'}\n`;
            }
        }
        await ctx.reply(message);
    } catch (error) {
        console.error('Products error:', error.message);
        await ctx.reply('⚠️ مشكلة في جلب المنتجات');
    }
});

// أمر إلغاء الطلب
bot.command('cancel', async (ctx) => {
    const userId = ctx.from.id;
    userSessions[userId] = { step: 'idle' };
    await ctx.reply('❌ تم إلغاء الطلب. تقدر تطلب تاني في أي وقت.');
});

// معالج الرسائل
bot.on('text', async (ctx) => {
    const userId = ctx.from.id;
    const userMessage = ctx.message.text;
    
    // تجاهل الأوامر
    if (userMessage.startsWith('/')) return;
    
    // إنشاء جلسة جديدة إذا لم توجد
    if (!userSessions[userId]) {
        userSessions[userId] = { step: 'idle' };
    }
    
    const session = userSessions[userId];
    
    try {
        await ctx.sendChatAction('typing');
        
        // حالة 1: بانتظار الطلب
        if (session.step === 'idle') {
            const doc = await getDoc();
            const productSheet = doc.sheetsByTitle['Products'];
            await productSheet.loadHeaderRow();
            const rows = await productSheet.getRows();
            
            // بناء قائمة المنتجات
            let productsText = '';
            for (const row of rows) {
                const product = row.get('المنتج');
                const price = row.get('السعر');
                const unit = row.get('الوحدة');
                if (product && price) {
                    productsText += `${product} (${price} جنيه/${unit || 'قطعة'})، `;
                }
            }
            
            // الاتصال بـ Groq (باستخدام الموديل الجديد)
            try {
                const aiResponse = await axios.post('https://api.groq.com/openai/v1/chat/completions', {
                    model: "llama-3.3-70b-versatile",
                    messages: [
                        {
                            role: "system",
                            content: `أنت موظف استقبال في سوبر ماركت مصري.
                            
المنتجات المتاحة: ${productsText}

مهمتك:
1. رد بالمصري (يا فندم، تؤمر)
2. إذا طلب العميل منتج موجود، احسب السعر
3. بعد حساب السعر، اطلب منه اسمه
4. استخدم عبارة "قولي اسمك الأول"`

                        },
                        { role: "user", content: userMessage }
                    ]
                }, {
                    headers: {
                        'Authorization': `Bearer ${process.env.GROQ_API_KEY}`,
                        'Content-Type': 'application/json'
                    },
                    timeout: 15000
                });
                
                const botReply = aiResponse.data.choices[0].message.content;
                await ctx.reply(botReply);
                
                // حفظ الطلب المؤقت
                session.orderText = userMessage;
                
                // استخراج السعر
                const priceMatch = botReply.match(/(\d+)\s*جنيه/);
                if (priceMatch) {
                    session.totalPrice = priceMatch[1];
                }
                
                // إذا طلب البوت الاسم، غير الحالة
                if (botReply.includes('اسم') || botReply.includes('قولي')) {
                    session.step = 'waiting_name';
                }
                
            } catch (aiError) {
                console.error('Groq error:', aiError.response?.data || aiError.message);
                await ctx.reply('عذرا يا فندم، فيه مشكلة في الاتصال. جرب تاني كدة.');
            }
        }
        // حالة 2: انتظار الاسم
        else if (session.step === 'waiting_name') {
            session.customerName = userMessage;
            session.step = 'waiting_address';
            await ctx.reply(`تسلم يا ${userMessage} 🙏\n\nدلوقتي قولي عنوانك (الشارع والمنطقة):`);
        }
        // حالة 3: انتظار العنوان
        else if (session.step === 'waiting_address') {
            session.customerAddress = userMessage;
            session.step = 'waiting_phone';
            await ctx.reply('تمام يا فندم 📍\n\nآخر حاجة، رقم تليفونك (مثال: 01001234567):');
        }
        // حالة 4: انتظار رقم الهاتف وحفظ الطلب
        else if (session.step === 'waiting_phone') {
            const phoneMatch = userMessage.match(/01[0125][0-9]{8}/);
            
            if (!phoneMatch) {
                await ctx.reply('❌ الرقم مش صحيح. أكتب رقم مصري 11 رقم يبدأ ب 01 (مثال: 01001234567)');
                return;
            }
            
            // حفظ الطلب في Google Sheets
            const doc = await getDoc();
            const orderSheet = doc.sheetsByTitle['Orders'];
            
            await orderSheet.addRow({
                'ID': userId.toString(),
                'اسم العميل': session.customerName,
                'الطلبات': session.orderText || 'غير محدد',
                'الإجمالي': session.totalPrice || '0',
                'العنوان': session.customerAddress,
                'رقم الهاتف': phoneMatch[0],
                'التاريخ': new Date().toLocaleString('ar-EG')
            });
            
            // إشعار للإدارة
            if (process.env.ADMIN_CHAT_ID) {
                await ctx.telegram.sendMessage(process.env.ADMIN_CHAT_ID,
                    `🛒 طلب جديد!\n👤 ${session.customerName}\n📦 ${session.orderText}\n💰 ${session.totalPrice} جنيه\n📍 ${session.customerAddress}\n📞 ${phoneMatch[0]}`
                ).catch(() => {});
            }
            
            await ctx.reply(`✅ تم تسجيل طلبك يا ${session.customerName}!\n\n📋 الطلب: ${session.orderText}\n💰 الإجمالي: ${session.totalPrice} جنيه\n📍 العنوان: ${session.customerAddress}\n\nشكراً لتسوقك معانا 🙏`);
            
            // إعادة ضبط الجلسة
            session.step = 'idle';
            delete session.customerName;
            delete session.customerAddress;
            delete session.orderText;
            delete session.totalPrice;
        }
        
    } catch (error) {
        console.error('Error:', error.message);
        await ctx.reply('⚠️ معلش فيه مشكلة تقنية. جرب تاني أو ابعت /cancel وابدأ من جديد.');
    }
});

// تصدير لـ Vercel
export default async (req, res) => {
    if (req.method === 'POST') {
        try {
            await bot.handleUpdate(req.body);
            res.status(200).json({ status: 'ok' });
        } catch (err) {
            console.error('Webhook error:', err);
            res.status(500).json({ error: err.message });
        }
    } else {
        res.status(200).send('✅ Bot is running!');
    }
};

// للتشغيل المحلي
if (process.env.NODE_ENV === 'development') {
    bot.launch();
    console.log('🤖 Bot running locally...');
}