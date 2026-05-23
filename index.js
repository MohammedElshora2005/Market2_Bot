import { Telegraf, Markup } from 'telegraf';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

const bot = new Telegraf(process.env.BOT_TOKEN); // تأكد من اسم المتغير في Vercel

const ADMIN_ID = process.env.ADMIN_CHAT_ID;
const ITEMS_PER_PAGE = 6;
const CACHE_TTL = 300000;

// ملاحظة: الـ Maps دي هتشتغل طالما الـ Serverless Instance صاحية، بس الأفضل مستقبلاً نربطها بـ Database أو شيت مخصص
const sessions = new Map();
const carts = new Map();
const offeredCustomers = new Map();
let currentPage = 0;
let currentProducts = [];
let currentSearch = '';
let lastOffersHash = '';
let isBroadcasting = false;

let productsCache = null;
let productsCacheTime = 0;
let offersCache = null;
let offersCacheTime = 0;

class UserSession {
    constructor(userId, name = '') {
        this.userId = userId;
        this.name = name;
        this.address = '';
        this.phone = '';
        this.status = 'main';
        this.lastActivity = Date.now();
        this.totalOrders = 0;
        this.totalSpent = 0;
        this.loyaltyPoints = 0;
        this.lastOrderNumber = null;
        this.editingOrder = null;
        this.editCart = [];
        this.pendingRating = null;
        this.tempOrderNumber = null;
    }
}

class CartItem {
    constructor(name, price, quantity = 1, unit = 'قطعة') {
        this.id = Date.now().toString();
        this.name = name;
        this.price = price;
        this.quantity = Math.min(quantity, 99);
        this.unit = unit;
        this.addedAt = Date.now();
    }
    get total() {
        return this.price * this.quantity;
    }
}

let docInstance = null;
let authInstance = null;

async function getAuth() {
    if (authInstance) return authInstance;
    const rawKey = process.env.GOOGLE_PRIVATE_KEY || '';
    const formattedKey = rawKey.replace(/\\n/g, '\n').replace(/"/g, '');
    authInstance = new JWT({
        email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        key: formattedKey,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    return authInstance;
}

async function getDoc() {
    if (docInstance) return docInstance;
    const auth = await getAuth();
    docInstance = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth); // تم تعديلها لـ SPREADSHEET_ID لتطابق مشروعك السابق
    await docInstance.loadInfo();
    return docInstance;
}

async function getProducts(forceRefresh = false) {
    const now = Date.now();
    if (!forceRefresh && productsCache && (now - productsCacheTime) < CACHE_TTL) {
        return productsCache;
    }
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['Products'];
        if (!sheet) return [];
        const rows = await sheet.getRows();
        const products = [];
        for (const row of rows) {
            const name = row.get('المنتج');
            const price = parseFloat(row.get('السعر'));
            const unit = row.get('الوحدة') || 'قطعة';
            const available = row.get('التوفر');
            let isAvailable = true;
            if (available === '0' || available === 'غير متوفر' || available === 'لا') {
                isAvailable = false;
            }
            if (name && price && !isNaN(price) && price > 0 && isAvailable) {
                products.push({ name, price, unit });
            }
        }
        productsCache = products;
        productsCacheTime = now;
        return products;
    } catch (error) {
        console.error(error);
        return productsCache || [];
    }
}

async function getOffers(forceRefresh = false) {
    const now = Date.now();
    if (!forceRefresh && offersCache && (now - offersCacheTime) < CACHE_TTL) {
        return offersCache;
    }
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['Offers'];
        if (!sheet) return [];
        const rows = await sheet.getRows();
        const offers = [];
        for (const row of rows) {
            const active = row.get('نشط');
            const text = row.get('العرض');
            const date = row.get('التاريخ');
            if (active === 'نعم' && text) {
                offers.push({ text, date: date || new Date().toLocaleString('ar-EG') });
            }
        }
        offersCache = offers;
        offersCacheTime = now;
        return offers;
    } catch (error) {
        return offersCache || [];
    }
}

async function saveOrder(userId, customerName, address, phone, items, total) {
    try {
        const doc = await getDoc();
        let sheet = doc.sheetsByTitle['Orders'];
        const orderText = items.map(item => `${item.quantity} × ${item.name} (${item.price}ج)`).join('\n• ');
        const timestamp = Date.now();
        const orderNumber = `ORD-${timestamp}-${userId.toString().slice(-4)}`;
        
        await sheet.addRow({
            'ID': userId.toString(),
            'اسم العميل': customerName,
            'الطلبات': `• ${orderText}`,
            'الإجمالي': total,
            'العنوان': address,
            'رقم الهاتف': phone,
            'التاريخ': new Date().toLocaleString('ar-EG'),
            'عدد العناصر': items.reduce((sum, i) => sum + i.quantity, 0),
            'رقم الطلب': orderNumber,
            'الحالة': 'جاري التنفيذ'
        });
        return { success: true, orderNumber };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

async function getUserOrders(userId) {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['Orders'];
        if (!sheet) return [];
        const rows = await sheet.getRows();
        const orders = [];
        for (const row of rows) {
            if (row.get('ID') === userId.toString()) {
                orders.push({
                    text: row.get('الطلبات'),
                    price: row.get('الإجمالي'),
                    date: row.get('التاريخ'),
                    status: row.get('الحالة') || 'جاري التنفيذ',
                    orderNumber: row.get('رقم الطلب') || '',
                    itemsText: row.get('الطلبات')
                });
            }
        }
        return orders.reverse();
    } catch (error) {
        return [];
    }
}

// القائمة الرئيسية للتفاعل
const mainKeyboard = Markup.inlineKeyboard([
    [Markup.button.callback('🛒 المنتجات', 'browse_products')],
    [Markup.button.callback('🛍️ السلة', 'view_cart')],
    [Markup.button.callback('🎁 العروض', 'view_offers')],
    [Markup.button.callback('📋 طلباتي', 'my_orders')]
]);

bot.start((ctx) => {
    ctx.reply('👋 أهلاً بك في سوبر ماركت بوب مارت الحوت! اختر من القائمة:', mainKeyboard);
});

bot.action('browse_products', async (ctx) => {
    await ctx.answerCbQuery();
    const products = await getProducts();
    if(products.length === 0) return ctx.reply('⚠️ لا توجد منتجات حالياً.');
    
    const buttons = products.map(p => [Markup.button.callback(`➕ ${p.name} - ${p.price}ج`, `add_${p.name}`)]);
    buttons.push([Markup.button.callback('🏠 الرئيسية', 'main_menu')]);
    ctx.reply('🛒 قائمة المنتجات المتاحة:', Markup.inlineKeyboard(buttons));
});

bot.action('main_menu', async (ctx) => {
    await ctx.answerCbQuery();
    ctx.reply('📱 القائمة الرئيسية:', mainKeyboard);
});

// ============ مسار الـ الـ Webhook النهائي لـ Vercel ============
module.exports = async (req, res) => {
    try {
        if (req.method === 'POST') {
            await bot.handleUpdate(req.body);
            res.status(200).send('OK');
        } else {
            res.status(200).send('سيرفر بوب مارت المطور يعمل بنجاح!');
        }
    } catch (error) {
        console.error(error);
        res.status(500).send('Internal Server Error');
    }
};
