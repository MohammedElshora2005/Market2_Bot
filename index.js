import { Telegraf, Markup } from 'telegraf';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

const bot = new Telegraf(process.env.BOT_TOKEN);

// ============ الإعدادات ============
const ADMIN_ID = process.env.ADMIN_CHAT_ID;
const ITEMS_PER_PAGE = 6;
const CACHE_TTL = 300000;
const CART_EXPIRY_HOURS = 48;

// ============ قاعدة البيانات المؤقتة ============
const sessions = new Map();      // جلسات المستخدمين
const carts = new Map();          // سلات الشراء
const userCooldowns = new Map();  // منع السبام
let currentPage = 0;
let currentProducts = [];
let currentSearch = '';

// ============ كاش البيانات ============
let productsCache = null;
let productsCacheTime = 0;
let offersCache = null;
let offersCacheTime = 0;

// ============ كلاس جلسة المستخدم ============
class UserSession {
    constructor(userId, name = '', username = '') {
        this.userId = userId;
        this.name = name;
        this.username = username;
        this.address = '';
        this.phone = '';
        this.status = 'main';  // main, ordering_name, ordering_address, ordering_phone, searching, tracking, feedback_note, editing
        this.lastActivity = Date.now();
        this.totalSpent = 0;
        this.loyaltyPoints = 0;
        this.lastOrderNumber = null;
        this.editingOrder = null;      // رقم الطلب الجاري تعديله
        this.editCart = [];             // منتجات التعديل المؤقتة
        this.pendingRating = null;      // تقييم معلق
        this.tempOrderNumber = null;    // رقم طلب مؤقت للإلغاء
    }
    
    updateActivity() {
        this.lastActivity = Date.now();
    }
    
    // حساب المستوى بناءً على إجمالي المشتريات
    getTier() {
        if (this.totalSpent >= 5000) return { name: '💎 بلاتينيوم', discount: 10, minSpent: 5000 };
        if (this.totalSpent >= 2000) return { name: '🌟 ذهبي', discount: 7, minSpent: 2000 };
        if (this.totalSpent >= 500) return { name: '⭐ فضي', discount: 5, minSpent: 500 };
        if (this.totalSpent >= 100) return { name: '🟤 برونزي', discount: 3, minSpent: 100 };
        return { name: '🟣 جديد', discount: 0, minSpent: 0 };
    }
    
    // إضافة نقاط ولاء
    addLoyaltyPoints(amount) {
        const points = Math.floor(amount / 10);
        this.loyaltyPoints += points;
        return points;
    }
}

// ============ كلاس عنصر السلة ============
class CartItem {
    constructor(name, price, quantity = 1) {
        this.id = Date.now().toString();
        this.name = name;
        this.price = price;
        this.quantity = Math.min(quantity, 99);
        this.addedAt = Date.now();
    }
    
    get total() {
        return this.price * this.quantity;
    }
}

// ============ دوال جوجل شيت ============
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
    docInstance = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
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
            const available = row.get('التوفر');
            let isAvailable = true;
            if (available === '0' || available === 'غير متوفر' || available === 'لا') {
                isAvailable = false;
            }
            if (name && price && !isNaN(price) && price > 0 && isAvailable) {
                products.push({ name, price });
            }
        }
        productsCache = products;
        productsCacheTime = now;
        console.log(`✅ تم تحميل ${products.length} منتج`);
        return products;
    } catch (error) {
        console.error('❌ خطأ في المنتجات:', error);
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
        console.error('❌ خطأ في العروض:', error);
        return offersCache || [];
    }
}

async function saveOrder(userId, customerName, address, phone, items, total) {
    try {
        const doc = await getDoc();
        let sheet = doc.sheetsByTitle['Orders'];
        if (!sheet) {
            sheet = await doc.addSheet({
                title: 'Orders',
                headerValues: ['ID', 'اسم العميل', 'الطلبات', 'الإجمالي', 'العنوان', 'رقم الهاتف', 'التاريخ', 'عدد العناصر', 'رقم الطلب', 'الحالة']
            });
        }
        
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
        console.error('❌ خطأ في حفظ الطلب:', error);
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
                    orderNumber: row.get('رقم الطلب') || '',
                    price: row.get('الإجمالي'),
                    date: row.get('التاريخ'),
                    status: row.get('الحالة') || 'جاري التنفيذ',
                    itemsText: row.get('الطلبات')
                });
            }
        }
        return orders.reverse();
    } catch (error) {
        console.error('❌ خطأ في جلب الطلبات:', error);
        return [];
    }
}

async function getOrderByNumber(orderNumber) {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['Orders'];
        if (!sheet) return null;
        const rows = await sheet.getRows();
        for (const row of rows) {
            if (row.get('رقم الطلب') === orderNumber) {
                const status = row.get('الحالة') || 'جاري التنفيذ';
                return {
                    orderNumber: row.get('رقم الطلب'),
                    name: row.get('اسم العميل'),
                    text: row.get('الطلبات'),
                    price: row.get('الإجمالي'),
                    address: row.get('العنوان'),
                    phone: row.get('رقم الهاتف'),
                    date: row.get('التاريخ'),
                    status: status,
                    canEdit: status === 'جاري التنفيذ',  // فقط "جاري التنفيذ" يمكن تعديله
                    row: row
                };
            }
        }
        return null;
    } catch (error) {
        console.error('❌ خطأ في جلب الطلب:', error);
        return null;
    }
}

async function updateOrderStatus(orderNumber, newStatus) {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['Orders'];
        if (!sheet) return false;
        const rows = await sheet.getRows();
        for (const row of rows) {
            if (row.get('رقم الطلب') === orderNumber) {
                row.set('الحالة', newStatus);
                await row.save();
                return true;
            }
        }
        return false;
    } catch (error) {
        console.error('❌ خطأ في تحديث الحالة:', error);
        return false;
    }
}

async function addItemsToOrder(orderNumber, newItems) {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['Orders'];
        if (!sheet) return false;
        const rows = await sheet.getRows();
        for (const row of rows) {
            if (row.get('رقم الطلب') === orderNumber) {
                const oldText = row.get('الطلبات');
                const newItemsText = newItems.map(item => `${item.quantity} × ${item.name} (${item.price}ج)`).join('\n• ');
                row.set('الطلبات', oldText + '\n• ' + newItemsText);
                
                const oldTotal = parseFloat(row.get('الإجمالي')) || 0;
                const newTotal = newItems.reduce((sum, item) => sum + (item.price * item.quantity), 0);
                row.set('الإجمالي', oldTotal + newTotal);
                
                const oldCount = parseInt(row.get('عدد العناصر')) || 0;
                const newCount = newItems.reduce((sum, item) => sum + item.quantity, 0);
                row.set('عدد العناصر', oldCount + newCount);
                
                await row.save();
                return true;
            }
        }
        return false;
    } catch (error) {
        console.error('❌ خطأ في إضافة منتجات:', error);
        return false;
    }
}

async function saveFeedback(userId, userName, rating, message, orderNumber = '') {
    try {
        const doc = await getDoc();
        let sheet = doc.sheetsByTitle['Feedbakes'];
        if (!sheet) {
            sheet = await doc.addSheet({
                title: 'Feedbakes',
                headerValues: ['ID', 'الاسم', 'التقييم', 'الرسالة', 'التاريخ', 'رقم الطلب', 'تم الرد']
            });
        }
        await sheet.addRow({
            'ID': userId.toString(),
            'الاسم': userName,
            'التقييم': rating,
            'الرسالة': message || '',
            'التاريخ': new Date().toLocaleString('ar-EG'),
            'رقم الطلب': orderNumber,
            'تم الرد': 'لا'
        });
        console.log(`✅ تم حفظ تقييم ${rating} نجوم من ${userName}`);
        return true;
    } catch (error) {
        console.error('❌ خطأ في حفظ التقييم:', error);
        return false;
    }
}

async function getUserTotalSpent(userId) {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['Orders'];
        if (!sheet) return 0;
        const rows = await sheet.getRows();
        let total = 0;
        for (const row of rows) {
            if (row.get('ID') === userId.toString()) {
                total += parseFloat(row.get('الإجمالي')) || 0;
            }
        }
        return total;
    } catch (error) {
        return 0;
    }
}

// ============ عرض المنتجات ============
async function showProducts(ctx, page = 0, searchQuery = '') {
    try {
        let products = await getProducts();
        if (products.length === 0) {
            await ctx.reply('⚠️ لا توجد منتجات متاحة حالياً', {
                ...Markup.inlineKeyboard([[Markup.button.callback('🏠 الرئيسية', 'main_menu')]])
            });
            return;
        }
        
        if (searchQuery) {
            products = products.filter(p => p.name.toLowerCase().includes(searchQuery.toLowerCase()));
            if (products.length === 0) {
                await ctx.reply(`🔍 لا توجد نتائج لـ "${searchQuery}"`);
                return;
            }
        }
        
        currentProducts = products;
        currentPage = page;
        currentSearch = searchQuery;
        
        const totalPages = Math.ceil(products.length / ITEMS_PER_PAGE);
        const start = page * ITEMS_PER_PAGE;
        const end = start + ITEMS_PER_PAGE;
        const pageProducts = products.slice(start, end);
        
        const buttons = [];
        for (const product of pageProducts) {
            buttons.push([
                Markup.button.callback(`➕ ${product.name} - ${product.price} ج`, `add_${product.name}_${product.price}`)
            ]);
        }
        
        const navRow = [];
        if (page > 0) navRow.push(Markup.button.callback('⬅️ السابق', 'prev_page'));
        navRow.push(Markup.button.callback(`📄 ${page + 1}/${totalPages}`, 'noop'));
        if (page < totalPages - 1) navRow.push(Markup.button.callback('التالي ➡️', 'next_page'));
        if (navRow.length > 1) buttons.push(navRow);
        
        buttons.push([Markup.button.callback('🔍 بحث', 'search_products')]);
        buttons.push([Markup.button.callback('🛍️ السلة', 'view_cart')]);
        buttons.push([Markup.button.callback('🏠 الرئيسية', 'main_menu')]);
        
        let msg = '🛒 <b>المنتجات</b>\n━━━━━━━━━━━━━━━\n';
        if (searchQuery) msg += `🔍 بحث: "${searchQuery}"\n`;
        msg += `📦 ${products.length} منتج\n━━━━━━━━━━━━━━━\n\n➕ اضغط على المنتج للإضافة`;
        
        await ctx.reply(msg, { parse_mode: 'HTML', ...Markup.inlineKeyboard(buttons) });
    } catch (error) {
        console.error('❌ خطأ في عرض المنتجات:', error);
        await ctx.reply('❌ حدث خطأ في عرض المنتجات');
    }
}

// ============ عرض السلة ============
async function showCart(ctx) {
    const userId = ctx.from.id;
    let cart = carts.get(userId) || [];
    
    if (cart.length === 0) {
        await ctx.reply('🛍️ سلتك فارغة', {
            ...Markup.inlineKeyboard([
                [Markup.button.callback('🛒 المنتجات', 'browse_products')],
                [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
            ])
        });
        return;
    }
    
    let total = 0;
    let msg = '🛍️ <b>سلة المشتريات</b>\n━━━━━━━━━━━━━━━\n\n';
    const buttons = [];
    
    for (let i = 0; i < cart.length; i++) {
        const item = cart[i];
        total += item.total;
        msg += `${i+1}. ${item.name}: ${item.quantity} × ${item.price} = ${item.total} ج\n`;
        buttons.push([
            Markup.button.callback('➕', `inc_${i}`),
            Markup.button.callback(`${item.quantity}`, 'noop'),
            Markup.button.callback('➖', `dec_${i}`),
            Markup.button.callback('❌', `rem_${i}`)
        ]);
    }
    msg += `\n━━━━━━━━━━━━━━━\n💰 <b>الإجمالي:</b> ${total} ج`;
    
    buttons.push([Markup.button.callback('➕ إضافة منتجات', 'browse_products')]);
    buttons.push([Markup.button.callback('✅ تأكيد الطلب', 'checkout')]);
    buttons.push([Markup.button.callback('🗑️ تفريغ السلة', 'clear_cart')]);
    buttons.push([Markup.button.callback('🏠 الرئيسية', 'main_menu')]);
    
    await ctx.reply(msg, { parse_mode: 'HTML', ...Markup.inlineKeyboard(buttons) });
}

// ============ عرض الطلبات ============
async function showOrders(ctx) {
    const userId = ctx.from.id;
    const orders = await getUserOrders(userId);
    
    if (orders.length === 0) {
        await ctx.reply('📭 لا يوجد طلبات سابقة', {
            ...Markup.inlineKeyboard([
                [Markup.button.callback('🛒 المنتجات', 'browse_products')],
                [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
            ])
        });
        return;
    }
    
    let msg = '📋 <b>طلباتي</b>\n━━━━━━━━━━━━━━━\n\n';
    const buttons = [];
    
    for (let i = 0; i < Math.min(orders.length, 10); i++) {
        const order = orders[i];
        let icon = order.status === 'تم التسليم' ? '✅' : order.status === 'في الطريق' ? '🚚' : order.status === 'ملغي' ? '❌' : '🟡';
        
        msg += `${i+1}. ${icon} <b>${order.date}</b>\n`;
        msg += `   🏷️ ${order.orderNumber}\n`;
        msg += `   💰 ${order.price} ج\n`;
        msg += `   📍 ${order.status}\n`;
        msg += `━━━━━━━━━━━━━━━\n\n`;
        buttons.push([Markup.button.callback(`🔍 تتبع الطلب`, `track_${order.orderNumber}`)]);
    }
    
    buttons.push([Markup.button.callback('📊 إحصائياتي', 'my_stats')]);
    buttons.push([Markup.button.callback('🏠 الرئيسية', 'main_menu')]);
    
    await ctx.reply(msg, { parse_mode: 'HTML', ...Markup.inlineKeyboard(buttons) });
}

// ============ عرض العروض ============
async function showOffers(ctx) {
    const offers = await getOffers();
    if (offers.length === 0) {
        await ctx.reply('🎁 لا توجد عروض حالياً', {
            ...Markup.inlineKeyboard([
                [Markup.button.callback('🛒 المنتجات', 'browse_products')],
                [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
            ])
        });
        return;
    }
    
    let msg = '🎁 <b>عروض بوب مارت</b>\n━━━━━━━━━━━━━━━\n\n';
    for (const offer of offers) {
        msg += `✨ ${offer.text}\n📅 ${offer.date}\n━━━━━━━━━━━━━━━\n`;
    }
    
    await ctx.reply(msg, {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard([
            [Markup.button.callback('🛒 المنتجات', 'browse_products')],
            [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
        ])
    });
}

// ============ تتبع الطلب (مع التحكم في أزرار التعديل حسب الحالة) ============
async function trackOrder(ctx, orderNumber) {
    if (!orderNumber || orderNumber.length < 10) {
        await ctx.reply('❌ رقم الطلب غير صحيح\n\nالرجاء إرسال رقم صحيح (مثال: ORD-1734567890123-7573)');
        return;
    }
    
    const order = await getOrderByNumber(orderNumber);
    if (!order) {
        await ctx.reply(`❌ لم يتم العثور على الطلب: ${orderNumber}`);
        return;
    }
    
    let icon = '🟡', statusText = '';
    if (order.status === 'جاري التنفيذ') {
        icon = '🟡';
        statusText = '✅ جاري تجهيز طلبك - يمكنك تعديل أو إلغاء الطلب';
    } else if (order.status === 'تم التجهيز') {
        icon = '📦';
        statusText = 'طلبك جاهز للتوصيل - لا يمكن التعديل الآن';
    } else if (order.status === 'في الطريق') {
        icon = '🚚';
        statusText = 'مندوب التوصيل في طريقه إليك - لا يمكن التعديل';
    } else if (order.status === 'تم التسليم') {
        icon = '✅';
        statusText = 'تم توصيل طلبك بنجاح - شكراً لتسوقك معنا';
    } else if (order.status === 'ملغي') {
        icon = '❌';
        statusText = 'تم إلغاء الطلب';
    }
    
    let msg = `📦 <b>تتبع الطلب</b>\n━━━━━━━━━━━━━━━\n\n`;
    msg += `🏷️ <b>رقم الطلب:</b> ${order.orderNumber}\n`;
    msg += `${icon} <b>الحالة:</b> ${order.status}\n`;
    msg += `📝 ${statusText}\n\n`;
    msg += `👤 <b>العميل:</b> ${order.name}\n`;
    msg += `💰 <b>الإجمالي:</b> ${order.price} ج\n`;
    msg += `📍 <b>العنوان:</b> ${order.address}\n`;
    msg += `📅 <b>التاريخ:</b> ${order.date}\n`;
    msg += `━━━━━━━━━━━━━━━\n\n`;
    msg += `📦 <b>المنتجات:</b>\n${order.text}`;
    
    const buttons = [];
    
    // ✅ فقط إذا كانت الحالة "جاري التنفيذ" تظهر أزرار التعديل والإلغاء
    if (order.canEdit) {
        buttons.push([Markup.button.callback('✏️ تعديل الطلب', `edit_order_${order.orderNumber}`)]);
        buttons.push([Markup.button.callback('❌ إلغاء الطلب', `cancel_order_${order.orderNumber}`)]);
    }
    
    // ✅ إذا كانت الحالة "تم التسليم" تظهر زر التقييم
    if (order.status === 'تم التسليم') {
        buttons.push([Markup.button.callback('⭐ تقييم الطلب', `rate_order_${order.orderNumber}`)]);
    }
    
    buttons.push([Markup.button.callback('📋 كل طلباتي', 'my_orders')]);
    buttons.push([Markup.button.callback('🏠 الرئيسية', 'main_menu')]);
    
    await ctx.reply(msg, { parse_mode: 'HTML', ...Markup.inlineKeyboard(buttons) });
}

// ============ عرض إحصائيات المستخدم ============
async function showStats(ctx) {
    const userId = ctx.from.id;
    const orders = await getUserOrders(userId);
    const session = sessions.get(userId) || new UserSession(userId);
    
    // حساب إجمالي المشتريات من قاعدة البيانات
    const totalSpent = await getUserTotalSpent(userId);
    session.totalSpent = totalSpent;
    sessions.set(userId, session);
    
    const tier = session.getTier();
    const pending = orders.filter(o => o.status === 'جاري التنفيذ').length;
    const completed = orders.filter(o => o.status === 'تم التسليم').length;
    
    let nextTierText = '';
    if (tier.name === '🟣 جديد') nextTierText = `\n🎯 أنفق ${100 - totalSpent} ج أخرى للوصول إلى مستوى برونزي`;
    else if (tier.name === '🟤 برونزي') nextTierText = `\n🎯 أنفق ${500 - totalSpent} ج أخرى للوصول إلى مستوى فضي`;
    else if (tier.name === '⭐ فضي') nextTierText = `\n🎯 أنفق ${2000 - totalSpent} ج أخرى للوصول إلى مستوى ذهبي`;
    else if (tier.name === '🌟 ذهبي') nextTierText = `\n🎯 أنفق ${5000 - totalSpent} ج أخرى للوصول إلى مستوى بلاتينيوم`;
    else nextTierText = `\n🏆 أنت في أعلى مستوى!`;
    
    const msg = `📊 <b>إحصائياتي</b>\n━━━━━━━━━━━━━━━\n\n` +
        `👤 <b>الاسم:</b> ${session.name || 'غير مسجل'}\n` +
        `🏆 <b>المستوى:</b> ${tier.name} (خصم ${tier.discount}%)\n` +
        `📦 <b>عدد الطلبات:</b> ${orders.length}\n` +
        `✅ <b>تم التسليم:</b> ${completed}\n` +
        `🔄 <b>قيد التنفيذ:</b> ${pending}\n` +
        `💰 <b>إجمالي المشتريات:</b> ${totalSpent} ج\n` +
        `⭐ <b>نقاط الولاء:</b> ${session.loyaltyPoints}\n` +
        `${nextTierText}\n\n` +
        `💡 <b>ملاحظة:</b> كل 10 ج = 1 نقطة ولاء\n` +
        `كلما زادت مشترياتك، ارتفع مستواك وزاد الخصم!`;
    
    await ctx.reply(msg, {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard([
            [Markup.button.callback('📋 طلباتي', 'my_orders')],
            [Markup.button.callback('⭐ تقييم الخدمة', 'feedback')],
            [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
        ])
    });
}

// ============ التقييم ============
async function showFeedbackButtons(ctx, orderNumber = '') {
    const suffix = orderNumber ? `_${orderNumber}` : '';
    await ctx.reply(
        '⭐ <b>تقييم الخدمة</b>\n\nكيف تقيم تجربتك مع بوب مارت؟',
        {
            parse_mode: 'HTML',
            ...Markup.inlineKeyboard([
                [Markup.button.callback('😡 1 نجمة', `rate_1${suffix}`), Markup.button.callback('😐 2 نجمتين', `rate_2${suffix}`), Markup.button.callback('🙂 3 نجوم', `rate_3${suffix}`)],
                [Markup.button.callback('😊 4 نجوم', `rate_4${suffix}`), Markup.button.callback('🤩 5 نجوم', `rate_5${suffix}`)],
                [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
            ])
        }
    );
}

// ============ تعديل الطلب ============
async function handleEditOrder(ctx, orderNumber) {
    const order = await getOrderByNumber(orderNumber);
    
    if (!order || !order.canEdit) {
        await ctx.reply('❌ لا يمكن تعديل هذا الطلب حالياً لأنه لا يزال قيد التنفيذ أو تم إلغاؤه');
        return;
    }
    
    const session = sessions.get(ctx.from.id) || new UserSession(ctx.from.id);
    session.status = 'editing';
    session.editingOrder = orderNumber;
    session.editCart = [];
    sessions.set(ctx.from.id, session);
    
    await ctx.reply(`✏️ <b>تعديل الطلب #${orderNumber}</b>\n\n` +
        `📦 المنتجات الحالية:\n${order.text}\n\n` +
        `➕ أضف منتجات جديدة بالضغط على "🛒 المنتجات"\n` +
        `المنتجات الجديدة ستضاف إلى الطلب الحالي\n\n` +
        `بعد الانتهاء، اضغط "إنهاء التعديل"`,
        {
            parse_mode: 'HTML',
            ...Markup.inlineKeyboard([
                [Markup.button.callback('🛒 المنتجات', 'browse_products')],
                [Markup.button.callback('✅ إنهاء التعديل', 'finish_edit')],
                [Markup.button.callback('❌ إلغاء', 'cancel_edit')],
                [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
            ])
        }
    );
}

async function finishEdit(ctx) {
    const userId = ctx.from.id;
    const session = sessions.get(userId);
    
    if (!session || !session.editingOrder || session.status !== 'editing') {
        await ctx.reply('❌ لا يوجد تعديل نشط');
        return;
    }
    
    if (session.editCart.length === 0) {
        await ctx.reply('⚠️ لم تقم بإضافة أي منتجات جديدة\n\nلإضافة منتجات، اضغط على "🛒 المنتجات" ثم اختر المنتجات');
        return;
    }
    
    const success = await addItemsToOrder(session.editingOrder, session.editCart);
    
    if (success) {
        const addedItems = session.editCart.map(item => `${item.quantity} × ${item.name}`).join(', ');
        const orderNum = session.editingOrder;
        session.editCart = [];
        session.status = 'main';
        session.editingOrder = null;
        sessions.set(userId, session);
        
        await ctx.reply(`✅ تم تعديل الطلب #${orderNum} بنجاح!\n\n📦 تم إضافة: ${addedItems}`, {
            ...Markup.inlineKeyboard([[Markup.button.callback('🔍 تتبع الطلب', `track_${orderNum}`)]])
        });
    } else {
        await ctx.reply('❌ حدث خطأ في تعديل الطلب');
    }
}

async function cancelEdit(ctx) {
    const userId = ctx.from.id;
    const session = sessions.get(userId);
    
    if (session) {
        session.status = 'main';
        session.editingOrder = null;
        session.editCart = [];
        sessions.set(userId, session);
    }
    
    await ctx.reply('❌ تم إلغاء التعديل', {
        ...Markup.inlineKeyboard([[Markup.button.callback('🏠 الرئيسية', 'main_menu')]])
    });
}

// ============ إلغاء الطلب ============
async function handleCancelOrder(ctx, orderNumber) {
    const order = await getOrderByNumber(orderNumber);
    
    if (!order || !order.canEdit) {
        await ctx.reply('❌ لا يمكن إلغاء هذا الطلب حالياً لأنه لا يزال قيد التنفيذ أو تم إلغاؤه');
        return;
    }
    
    const session = sessions.get(ctx.from.id) || new UserSession(ctx.from.id);
    session.tempOrderNumber = orderNumber;
    sessions.set(ctx.from.id, session);
    
    await ctx.reply(`⚠️ <b>تأكيد إلغاء الطلب</b>\n\n` +
        `🏷️ ${order.orderNumber}\n` +
        `💰 ${order.price} ج\n\n` +
        `هل أنت متأكد من إلغاء هذا الطلب؟ لا يمكن التراجع!`,
        {
            parse_mode: 'HTML',
            ...Markup.inlineKeyboard([
                [Markup.button.callback('✅ نعم، ألغي', 'confirm_cancel')],
                [Markup.button.callback('❌ لا، عودة', `track_${orderNumber}`)]
            ])
        }
    );
}

async function confirmCancel(ctx) {
    const userId = ctx.from.id;
    const session = sessions.get(userId);
    
    if (!session || !session.tempOrderNumber) {
        await ctx.reply('❌ حدث خطأ');
        return;
    }
    
    const success = await updateOrderStatus(session.tempOrderNumber, 'ملغي');
    
    if (success) {
        await ctx.reply(`✅ تم إلغاء الطلب #${session.tempOrderNumber} بنجاح`, {
            ...Markup.inlineKeyboard([
                [Markup.button.callback('🛒 طلب جديد', 'browse_products')],
                [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
            ])
        });
    } else {
        await ctx.reply('❌ لا يمكن إلغاء هذا الطلب الآن');
    }
    
    session.tempOrderNumber = null;
    sessions.set(userId, session);
}

// ============ معالج الطلب ============
async function startCheckout(ctx) {
    const userId = ctx.from.id;
    let cart = carts.get(userId) || [];
    
    if (cart.length === 0) {
        await ctx.reply('🛍️ سلتك فارغة');
        return;
    }
    
    const total = cart.reduce((sum, item) => sum + item.total, 0);
    const session = sessions.get(userId) || new UserSession(userId);
    session.status = 'ordering_name';
    sessions.set(userId, session);
    
    const cartItems = cart.map((item, i) => `${i+1}. ${item.quantity} × ${item.name} = ${item.total} ج`).join('\n');
    
    await ctx.reply(`🛍️ <b>تأكيد الطلب</b>\n━━━━━━━━━━━━━━━\n\n${cartItems}\n\n💰 <b>الإجمالي:</b> ${total} ج\n━━━━━━━━━━━━━━━\n\n📝 أرسل اسمك الكامل:`, { parse_mode: 'HTML' });
}

// ============ البحث والتتبع ============
async function handleSearch(ctx) {
    await ctx.reply('🔍 اكتب اسم المنتج:', {
        ...Markup.inlineKeyboard([[Markup.button.callback('🏠 الرئيسية', 'main_menu')]])
    });
    const session = sessions.get(ctx.from.id) || new UserSession(ctx.from.id);
    session.status = 'searching';
    sessions.set(ctx.from.id, session);
}

async function handleTrack(ctx) {
    await ctx.reply('🔍 <b>تتبع الطلب</b>\n\n✏️ أرسل رقم الطلب (مثال: ORD-1734567890123-7573):\n\n📌 يمكنك نسخ الرقم من رسالة تأكيد الطلب', {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard([
            [Markup.button.callback('📋 طلباتي', 'my_orders')],
            [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
        ])
    });
    const session = sessions.get(ctx.from.id) || new UserSession(ctx.from.id);
    session.status = 'tracking';
    sessions.set(ctx.from.id, session);
}

// ============ القائمة الرئيسية ============
const mainKeyboard = Markup.inlineKeyboard([
    [Markup.button.callback('🛒 المنتجات', 'browse_products')],
    [Markup.button.callback('🛍️ السلة', 'view_cart')],
    [Markup.button.callback('🎁 العروض', 'view_offers')],
    [Markup.button.callback('📋 طلباتي', 'my_orders')],
    [Markup.button.callback('🔍 تتبع طلب', 'track_order')],
    [Markup.button.callback('📊 إحصائياتي', 'my_stats')],
    [Markup.button.callback('⭐ تقييم', 'feedback')],
    [Markup.button.callback('📞 الدعم', 'support')]
]);

// ============ معالج الأزرار ============

bot.start(async (ctx) => {
    const userId = ctx.from.id;
    let session = sessions.get(userId);
    if (!session) {
        session = new UserSession(userId, ctx.from.first_name, ctx.from.username);
        sessions.set(userId, session);
    }
    if (!carts.has(userId)) carts.set(userId, []);
    
    // جلب إجمالي المشتريات من قاعدة البيانات
    session.totalSpent = await getUserTotalSpent(userId);
    const tier = session.getTier();
    
    const welcomeMsg = `🌟 <b>مرحباً بك في بوب مارت!</b> 🌟\n━━━━━━━━━━━━━━━\n\n` +
        `👤 <b>الاسم:</b> ${ctx.from.first_name}\n` +
        `🏆 <b>مستواك:</b> ${tier.name} (خصم ${tier.discount}%)\n\n` +
        `🛒 أكبر سوبر ماركت في العالم!\n` +
        `✅ توصيل سريع - أسعار تنافسية - جودة عالية\n━━━━━━━━━━━━━━━\n\n` +
        `📌 اختر من القائمة:`;
    
    await ctx.reply(welcomeMsg, { parse_mode: 'HTML', ...mainKeyboard });
});

bot.action('main_menu', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await ctx.reply('🏠 القائمة الرئيسية', { ...mainKeyboard });
});

bot.action('browse_products', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await showProducts(ctx);
});

bot.action('view_cart', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await showCart(ctx);
});

bot.action('view_offers', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await showOffers(ctx);
});

bot.action('my_orders', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await showOrders(ctx);
});

bot.action('my_stats', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await showStats(ctx);
});

bot.action('track_order', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await handleTrack(ctx);
});

bot.action('feedback', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await showFeedbackButtons(ctx);
});

bot.action('support', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await ctx.reply('📞 <b>الدعم الفني</b>\n━━━━━━━━━━━━━━━\n\n' +
        '📱 واتساب: <code>01020063819</code>\n' +
        '✈️ تليجرام: @Muhamedhosny\n\n' +
        '⏰ أوقات العمل: 9 ص - 9 م', {
        parse_mode: 'HTML'
    });
});

bot.action('search_products', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await handleSearch(ctx);
});

bot.action('next_page', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await showProducts(ctx, currentPage + 1, currentSearch);
});

bot.action('prev_page', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await showProducts(ctx, currentPage - 1, currentSearch);
});

bot.action(/track_(.+)/, async (ctx) => {
    const orderNumber = ctx.match[1];
    await ctx.answerCbQuery().catch(() => {});
    await trackOrder(ctx, orderNumber);
});

bot.action('checkout', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await startCheckout(ctx);
});

bot.action('clear_cart', async (ctx) => {
    carts.set(ctx.from.id, []);
    await ctx.answerCbQuery().catch(() => {});
    await ctx.reply('🗑️ تم تفريغ السلة');
    await showCart(ctx);
});

bot.action(/inc_(\d+)/, async (ctx) => {
    const userId = ctx.from.id;
    const index = parseInt(ctx.match[1]);
    const cart = carts.get(userId) || [];
    if (cart[index] && cart[index].quantity < 99) {
        cart[index].quantity++;
        carts.set(userId, cart);
    }
    await showCart(ctx);
    await ctx.answerCbQuery().catch(() => {});
});

bot.action(/dec_(\d+)/, async (ctx) => {
    const userId = ctx.from.id;
    const index = parseInt(ctx.match[1]);
    const cart = carts.get(userId) || [];
    if (cart[index]) {
        if (cart[index].quantity > 1) {
            cart[index].quantity--;
        } else {
            cart.splice(index, 1);
        }
        carts.set(userId, cart);
        await showCart(ctx);
    }
    await ctx.answerCbQuery().catch(() => {});
});

bot.action(/rem_(\d+)/, async (ctx) => {
    const userId = ctx.from.id;
    const index = parseInt(ctx.match[1]);
    const cart = carts.get(userId) || [];
    if (cart[index]) {
        cart.splice(index, 1);
        carts.set(userId, cart);
        await showCart(ctx);
    }
    await ctx.answerCbQuery().catch(() => {});
});

bot.action(/add_(.+)_(.+)/, async (ctx) => {
    const productName = ctx.match[1];
    const productPrice = parseFloat(ctx.match[2]);
    const userId = ctx.from.id;
    const session = sessions.get(userId);
    
    await ctx.answerCbQuery(`✅ تم إضافة ${productName}`).catch(() => {});
    
    // ✅ التحقق: إذا كان المستخدم في وضع تعديل الطلب
    if (session && session.status === 'editing') {
        const existing = session.editCart.find(item => item.name === productName);
        if (existing) {
            existing.quantity++;
        } else {
            session.editCart.push({ name: productName, price: productPrice, quantity: 1 });
        }
        sessions.set(userId, session);
        await ctx.reply(`✅ تم إضافة ${productName} إلى تعديل الطلب\n\n📦 المنتجات المضافة حالياً: ${session.editCart.map(i => `${i.quantity}×${i.name}`).join(', ')}`);
    } 
    // ✅ الوضع العادي: إضافة إلى السلة
    else {
        let cart = carts.get(userId) || [];
        const existing = cart.find(item => item.name === productName);
        if (existing) {
            existing.quantity++;
        } else {
            cart.push(new CartItem(productName, productPrice, 1));
        }
        carts.set(userId, cart);
        await ctx.reply(`✅ تم إضافة ${productName} إلى السلة\n🛍️ عدد المنتجات في السلة: ${cart.length}`);
    }
});

bot.action('noop', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
});

// أزرار التعديل والإلغاء
bot.action(/edit_order_(.+)/, async (ctx) => {
    const orderNumber = ctx.match[1];
    await ctx.answerCbQuery().catch(() => {});
    await handleEditOrder(ctx, orderNumber);
});

bot.action(/cancel_order_(.+)/, async (ctx) => {
    const orderNumber = ctx.match[1];
    await ctx.answerCbQuery().catch(() => {});
    await handleCancelOrder(ctx, orderNumber);
});

bot.action('finish_edit', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await finishEdit(ctx);
});

bot.action('cancel_edit', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await cancelEdit(ctx);
});

bot.action('confirm_cancel', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await confirmCancel(ctx);
});

// أزرار التقييم
for (let i = 1; i <= 5; i++) {
    bot.action(new RegExp(`rate_${i}(?:_(.*))?`), async (ctx) => {
        const rating = i;
        const orderNumber = ctx.match[1] || '';
        const userId = ctx.from.id;
        const session = sessions.get(userId) || new UserSession(userId);
        
        await ctx.answerCbQuery(`✅ شكراً لتقييمك ${rating} نجوم!`).catch(() => {});
        
        // ✅ حفظ التقييم في Google Sheets
        const saved = await saveFeedback(userId, session.name || ctx.from.first_name, rating, '', orderNumber);
        
        if (saved) {
            await ctx.reply(
                `🙏 <b>شكراً لتقييمك ${'⭐'.repeat(rating)}</b>\n\n` +
                `تقييمك مهم جداً لتطوير خدماتنا.\n` +
                `هل تريد إضافة ملاحظات إضافية؟`,
                {
                    parse_mode: 'HTML',
                    ...Markup.inlineKeyboard([
                        [Markup.button.callback('📝 إضافة ملاحظة', `add_note_${orderNumber}_${rating}`)],
                        [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
                    ])
                }
            );
        } else {
            await ctx.reply('❌ حدث خطأ في حفظ التقييم، حاول مرة أخرى');
        }
    });
}

bot.action(/add_note_(.*)_(\d+)/, async (ctx) => {
    const orderNumber = ctx.match[1];
    const rating = ctx.match[2];
    await ctx.answerCbQuery().catch(() => {});
    
    const session = sessions.get(ctx.from.id) || new UserSession(ctx.from.id);
    session.status = 'feedback_note';
    session.pendingRating = { orderNumber, rating };
    sessions.set(ctx.from.id, session);
    
    await ctx.reply('📝 <b>أضف ملاحظاتك</b>\n\n✏️ اكتب ملاحظاتك أو اقتراحاتك هنا:');
});

bot.action(/rate_order_(.+)/, async (ctx) => {
    const orderNumber = ctx.match[1];
    await ctx.answerCbQuery().catch(() => {});
    
    const session = sessions.get(ctx.from.id) || new UserSession(ctx.from.id);
    session.pendingRating = orderNumber;
    sessions.set(ctx.from.id, session);
    
    await showFeedbackButtons(ctx, orderNumber);
});

// ============ معالج النصوص ============
bot.on('text', async (ctx) => {
    const userId = ctx.from.id;
    const text = ctx.message.text.trim();
    let session = sessions.get(userId);
    
    if (!session) {
        session = new UserSession(userId, ctx.from.first_name, ctx.from.username);
        sessions.set(userId, session);
    }
    
    // منع السبام
    const lastMessage = userCooldowns.get(userId);
    if (lastMessage && Date.now() - lastMessage < 1000) return;
    userCooldowns.set(userId, Date.now());
    
    // ✅ ملاحظات التقييم (هذا أولوية عالية)
    if (session.status === 'feedback_note' && session.pendingRating) {
        const { orderNumber, rating } = session.pendingRating;
        await saveFeedback(userId, session.name || ctx.from.first_name, parseInt(rating), text, orderNumber);
        session.status = 'main';
        session.pendingRating = null;
        sessions.set(userId, session);
        
        await ctx.reply('🙏 <b>شكراً لملاحظاتك!</b>\n\nتم استلام ملاحظاتك وسنعمل على تحسين خدماتنا.', {
            parse_mode: 'HTML',
            ...Markup.inlineKeyboard([[Markup.button.callback('🏠 الرئيسية', 'main_menu')]])
        });
        return;
    }
    
    // تتبع طلب
    if (session.status === 'tracking') {
        session.status = 'main';
        sessions.set(userId, session);
        await trackOrder(ctx, text);
        return;
    }
    
    // بحث
    if (session.status === 'searching') {
        session.status = 'main';
        sessions.set(userId, session);
        await showProducts(ctx, 0, text);
        return;
    }
    
    // إدخال اسم الطلب
    if (session.status === 'ordering_name') {
        if (text.length < 3) {
            await ctx.reply('❌ الاسم قصير جداً، أرسل اسمك الكامل:');
            return;
        }
        session.name = text;
        session.status = 'ordering_address';
        sessions.set(userId, session);
        await ctx.reply(`✅ تم حفظ: ${text}\n\n📍 أرسل عنوانك (الشارع، المنطقة، المدينة):`);
        return;
    }
    
    // إدخال عنوان
    if (session.status === 'ordering_address') {
        if (text.length < 10) {
            await ctx.reply('❌ العنوان غير مفصل، أرسل عنواناً مفصلاً:');
            return;
        }
        session.address = text;
        session.status = 'ordering_phone';
        sessions.set(userId, session);
        await ctx.reply(`✅ تم حفظ العنوان\n\n📞 أرسل رقم هاتفك (مثال: 01001234567):`);
        return;
    }
    
    // إدخال هاتف وحفظ الطلب
    if (session.status === 'ordering_phone') {
        const phoneMatch = text.match(/01[0125][0-9]{8}/);
        if (!phoneMatch) {
            await ctx.reply('❌ رقم غير صحيح، أرسل رقم مصري 11 رقم يبدأ بـ 01:');
            return;
        }
        
        session.phone = phoneMatch[0];
        const cart = carts.get(userId) || [];
        const total = cart.reduce((sum, item) => sum + item.total, 0);
        
        const saved = await saveOrder(userId, session.name, session.address, session.phone, cart, total);
        
        if (saved.success) {
            // تحديث إجمالي المشتريات ونقاط الولاء
            session.totalSpent += total;
            const pointsEarned = session.addLoyaltyPoints(total);
            const tier = session.getTier();
            
            carts.set(userId, []);
            session.status = 'main';
            sessions.set(userId, session);
            
            const successMessage = 
`✅ <b>تم تسجيل طلبك بنجاح!</b>
━━━━━━━━━━━━━━━

🏷️ <b>رقم الطلب:</b> <code>${saved.orderNumber}</code>

👤 <b>الاسم:</b> ${session.name}
💰 <b>الإجمالي:</b> ${total} ج
📍 <b>العنوان:</b> ${session.address}
📞 <b>الهاتف:</b> ${session.phone}

✨ <b>حصلت على ${pointsEarned} نقطة ولاء</b>
🏆 <b>مستواك الحالي:</b> ${tier.name} (خصم ${tier.discount}%)

📌 <b>لتتبع طلبك:</b>
اضغط على 🔍 تتبع طلب في القائمة وأرسل: <code>${saved.orderNumber}</code>

🙏 شكراً لتسوقك مع بوب مارت!`;
            
            await ctx.reply(successMessage, {
                parse_mode: 'HTML',
                ...Markup.inlineKeyboard([
                    [Markup.button.callback('🔍 تتبع هذا الطلب', `track_${saved.orderNumber}`)],
                    [Markup.button.callback('⭐ تقييم الخدمة', 'feedback')],
                    [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
                ])
            });
        } else {
            await ctx.reply('❌ حدث خطأ في حفظ الطلب، حاول مرة أخرى أو تواصل مع الدعم');
        }
        return;
    }
    
    // بحث عادي
    if (text.length > 2 && !text.startsWith('/')) {
        await showProducts(ctx, 0, text);
    } else if (!text.startsWith('/')) {
        await ctx.reply('🔍 استخدم الأزرار للتنقل', { ...mainKeyboard });
    }
});

// ============ تنظيف الجلسات ============
setInterval(() => {
    const now = Date.now();
    for (const [userId, session] of sessions.entries()) {
        if (now - session.lastActivity > 30 * 60 * 1000) {
            sessions.delete(userId);
        } else {
            session.updateActivity();
        }
    }
    for (const [userId, cart] of carts.entries()) {
        if (cart.length > 0) {
            const oldestItem = Math.min(...cart.map(i => i.addedAt));
            if (Date.now() - oldestItem > CART_EXPIRY_HOURS * 60 * 60 * 1000) {
                carts.delete(userId);
            }
        }
    }
}, 30 * 60 * 1000);

// ============ التحميل الأولي ============
setTimeout(async () => {
    console.log('🤖 جاري تحميل البوت...');
    await getProducts(true);
    console.log('✅ تم تحميل البوت بنجاح');
}, 5000);

// ============ مسار الـ Webhook لـ Vercel ============
export default async function handler(req, res) {
    try {
        if (req.method === 'POST') {
            await bot.handleUpdate(req.body);
            res.status(200).json({ status: 'ok' });
        } else {
            res.status(200).send('✅ Bob Mart Bot is running!');
        }
    } catch (error) {
        console.error('❌ Webhook error:', error);
        res.status(500).json({ error: error.message });
    }
}

// تشغيل محلي
if (process.env.NODE_ENV !== 'production') {
    bot.launch();
    console.log('🤖 Bob Mart Bot is running locally...');
}

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
