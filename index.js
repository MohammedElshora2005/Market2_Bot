import { Telegraf, Markup } from 'telegraf';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

const bot = new Telegraf(process.env.BOT_TOKEN);

// ============ الإعدادات المتقدمة ============
const ADMIN_ID = process.env.ADMIN_CHAT_ID;
const ITEMS_PER_PAGE = 6;
const CACHE_TTL = 300000;
const CART_EXPIRY_HOURS = 48;
const OFFERS_CHECK_INTERVAL = 6 * 60 * 60 * 1000;
const MAX_ORDERS_DISPLAY = 10;
const MAX_CART_ITEMS = 50;
const MIN_ORDER_AMOUNT = 10;
const DELIVERY_FEE = 0;
const FREE_DELIVERY_AMOUNT = 200;

// ============ قاعدة البيانات المؤقتة ============
const sessions = new Map();
const carts = new Map();
const userCooldowns = new Map();
const offeredCustomers = new Map();
const broadcastQueue = new Map();
let currentPage = 0;
let currentProducts = [];
let currentSearch = '';
let lastOffersHash = '';
let isBroadcasting = false;

// ============ كاش البيانات ============
let productsCache = null;
let productsCacheTime = 0;
let offersCache = null;
let offersCacheTime = 0;
let statsCache = null;
let statsCacheTime = 0;

// ============ كلاس جلسة المستخدم المتطور ============
class UserSession {
    constructor(userId, name = '', username = '') {
        this.userId = userId;
        this.name = name;
        this.username = username;
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
        this.language = 'ar';
        this.notifications = true;
        this.favoriteProducts = [];
        this.lastSearchQuery = '';
        this.referralCode = null;
        this.referredBy = null;
    }
    
    updateActivity() {
        this.lastActivity = Date.now();
    }
    
    addLoyaltyPoints(amount) {
        this.loyaltyPoints += Math.floor(amount / 10);
        return this.loyaltyPoints;
    }
    
    getTier() {
        if (this.totalSpent >= 5000) return { name: '💎 بلاتينيوم', discount: 10, color: '🔷' };
        if (this.totalSpent >= 2000) return { name: '🌟 ذهبي', discount: 7, color: '🟡' };
        if (this.totalSpent >= 500) return { name: '⭐ فضي', discount: 5, color: '⚪' };
        if (this.totalSpent >= 100) return { name: '🟢 برونزي', discount: 3, color: '🟤' };
        return { name: '🟣 جديد', discount: 0, color: '🟣' };
    }
}

// ============ كلاس عنصر السلة المتطور ============
class CartItem {
    constructor(name, price, quantity = 1, unit = 'قطعة') {
        this.id = Date.now().toString();
        this.name = name;
        this.price = price;
        this.quantity = Math.min(quantity, 99);
        this.unit = unit;
        this.addedAt = Date.now();
        this.notes = '';
    }
    
    get total() {
        return this.price * this.quantity;
    }
    
    get display() {
        return `${this.quantity} × ${this.name} (${this.price}ج) = ${this.total}ج`;
    }
}

// ============ دوال جوجل شيت المتقدمة ============
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
            const unit = row.get('الوحدة') || 'قطعة';
            const available = row.get('التوفر');
            const category = row.get('التصنيف') || 'عام';
            const description = row.get('الوصف') || '';
            let isAvailable = true;
            if (available === '0' || available === 'غير متوفر' || available === 'لا') {
                isAvailable = false;
            }
            if (name && price && !isNaN(price) && price > 0 && isAvailable) {
                products.push({ name, price, unit, category, description });
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
            const code = row.get('الكود') || '';
            const discount = row.get('الخصم') || 0;
            if (active === 'نعم' && text) {
                offers.push({ text, date: date || new Date().toLocaleString('ar-EG'), code, discount: parseInt(discount) });
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

async function saveOrder(userId, customerName, address, phone, items, total, discount = 0, finalTotal = null) {
    try {
        const doc = await getDoc();
        let sheet = doc.sheetsByTitle['Orders'];
        if (!sheet) {
            sheet = await doc.addSheet({
                title: 'Orders',
                headerValues: ['ID', 'اسم العميل', 'الطلبات', 'الإجمالي', 'الخصم', 'الصافي', 'العنوان', 'رقم الهاتف', 'التاريخ', 'عدد العناصر', 'رقم الطلب', 'الحالة', 'ملاحظات']
            });
        }
        
        const orderText = items.map(item => `${item.quantity} × ${item.name} (${item.price}ج)`).join('\n• ');
        const timestamp = Date.now();
        const orderNumber = `ORD-${timestamp}-${userId.toString().slice(-4)}`;
        const netTotal = finalTotal || (total - discount);
        
        await sheet.addRow({
            'ID': userId.toString(),
            'اسم العميل': customerName,
            'الطلبات': `• ${orderText}`,
            'الإجمالي': total,
            'الخصم': discount,
            'الصافي': netTotal,
            'العنوان': address,
            'رقم الهاتف': phone,
            'التاريخ': new Date().toLocaleString('ar-EG'),
            'عدد العناصر': items.reduce((sum, i) => sum + i.quantity, 0),
            'رقم الطلب': orderNumber,
            'الحالة': 'جاري التنفيذ',
            'ملاحظات': ''
        });
        
        return { success: true, orderNumber, netTotal };
    } catch (error) {
        console.error('❌ خطأ في حفظ الطلب:', error);
        return { success: false, error: error.message };
    }
}

async function getUserOrders(userId, limit = 50) {
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
                    netPrice: row.get('الصافي') || row.get('الإجمالي'),
                    discount: row.get('الخصم') || 0,
                    date: row.get('التاريخ'),
                    status: row.get('الحالة') || 'جاري التنفيذ',
                    orderNumber: row.get('رقم الطلب') || '',
                    itemsText: row.get('الطلبات'),
                    address: row.get('العنوان'),
                    phone: row.get('رقم الهاتف')
                });
            }
        }
        return orders.reverse().slice(0, limit);
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
                    netPrice: row.get('الصافي') || row.get('الإجمالي'),
                    discount: row.get('الخصم') || 0,
                    address: row.get('العنوان'),
                    phone: row.get('رقم الهاتف'),
                    date: row.get('التاريخ'),
                    status: status,
                    canEdit: status === 'جاري التنفيذ',
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

async function updateOrderStatus(orderNumber, newStatus, notes = '') {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['Orders'];
        if (!sheet) return false;
        const rows = await sheet.getRows();
        for (const row of rows) {
            if (row.get('رقم الطلب') === orderNumber) {
                row.set('الحالة', newStatus);
                if (notes) row.set('ملاحظات', notes);
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

async function getAllCustomers() {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['Orders'];
        if (!sheet) return [];
        const rows = await sheet.getRows();
        const customersMap = new Map();
        for (const row of rows) {
            const id = row.get('ID');
            const name = row.get('اسم العميل');
            const phone = row.get('رقم الهاتف');
            if (id && name && !customersMap.has(id)) {
                customersMap.set(id, { id, name, phone });
            }
        }
        return Array.from(customersMap.values());
    } catch (error) {
        console.error('❌ خطأ في جلب العملاء:', error);
        return [];
    }
}

async function saveFeedback(userId, userName, rating, message, orderNumber = '') {
    try {
        const doc = await getDoc();
        let sheet = doc.sheetsByTitle('Feedbakes');
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

async function getSystemStats() {
    const now = Date.now();
    if (statsCache && (now - statsCacheTime) < CACHE_TTL) return statsCache;
    
    try {
        const doc = await getDoc();
        const ordersSheet = doc.sheetsByTitle['Orders'];
        const feedbackSheet = doc.sheetsByTitle['Feedbakes'];
        
        let totalOrders = 0;
        let totalRevenue = 0;
        let pendingOrders = 0;
        
        if (ordersSheet) {
            const rows = await ordersSheet.getRows();
            totalOrders = rows.length;
            for (const row of rows) {
                totalRevenue += parseFloat(row.get('الصافي')) || parseFloat(row.get('الإجمالي')) || 0;
                if (row.get('الحالة') === 'جاري التنفيذ') pendingOrders++;
            }
        }
        
        let avgRating = 0;
        if (feedbackSheet) {
            const rows = await feedbackSheet.getRows();
            let totalRating = 0;
            for (const row of rows) {
                totalRating += parseInt(row.get('التقييم')) || 0;
            }
            avgRating = rows.length > 0 ? (totalRating / rows.length).toFixed(1) : 0;
        }
        
        statsCache = { totalOrders, totalRevenue, pendingOrders, avgRating };
        statsCacheTime = now;
        return statsCache;
    } catch (error) {
        console.error('❌ خطأ في جلب الإحصائيات:', error);
        return { totalOrders: 0, totalRevenue: 0, pendingOrders: 0, avgRating: 0 };
    }
}

// ============ دوال البث التلقائي للعروض ============
function getOffersHash(offers) {
    return JSON.stringify(offers.map(o => ({ text: o.text, date: o.date, code: o.code })));
}

async function broadcastOffersToAllCustomers() {
    if (isBroadcasting) return;
    try {
        isBroadcasting = true;
        const offers = await getOffers(true);
        const currentHash = getOffersHash(offers);
        if (currentHash === lastOffersHash || offers.length === 0) return;
        
        console.log('📢 بث العروض الجديدة...');
        const customers = await getAllCustomers();
        let sentCount = 0;
        
        for (const customer of customers) {
            const customerKey = `${customer.id}_${currentHash}`;
            if (offeredCustomers.has(customerKey)) continue;
            try {
                const message = `🎁 <b>عروض حصرية!</b> 🎁\n━━━━━━━━━━━━━━━\n\n` +
                    offers.map(o => `✨ ${o.text}\n📅 ${o.date}${o.code ? `\n🏷️ كود: ${o.code}` : ''}`).join('\n━━━━━━━━━━━━━━━\n') +
                    `\n\n⬇️ <b>اطلب الآن واستفد من العروض!</b>`;
                await bot.telegram.sendMessage(customer.id, message, {
                    parse_mode: 'HTML',
                    ...Markup.inlineKeyboard([
                        [Markup.button.callback('🛒 تصفح المنتجات', 'browse_products')],
                        [Markup.button.callback('🎁 عرض العروض', 'view_offers')],
                        [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
                    ])
                });
                offeredCustomers.set(customerKey, true);
                sentCount++;
                await new Promise(r => setTimeout(r, 100));
            } catch (err) { console.error(`فشل الإرسال للعميل ${customer.id}:`, err.message); }
        }
        console.log(`✅ تم بث العروض لـ ${sentCount} عميل`);
        lastOffersHash = currentHash;
        setTimeout(() => {
            for (const key of offeredCustomers.keys()) {
                if (key.includes(currentHash)) offeredCustomers.delete(key);
            }
        }, 24 * 60 * 60 * 1000);
    } finally { isBroadcasting = false; }
}

// ============ دوال عرض المنتجات المتقدمة ============
async function showProducts(ctx, page = 0, searchQuery = '', category = '') {
    try {
        let products = await getProducts();
        if (products.length === 0) {
            await ctx.reply('⚠️ لا توجد منتجات متاحة حالياً', {
                ...Markup.inlineKeyboard([[Markup.button.callback('🏠 الرئيسية', 'main_menu')]])
            });
            return;
        }
        
        if (searchQuery) {
            products = products.filter(p => p.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
                (p.description && p.description.toLowerCase().includes(searchQuery.toLowerCase())));
            if (products.length === 0) {
                await ctx.reply(`🔍 لا توجد نتائج لـ "${searchQuery}"`);
                return;
            }
        }
        
        if (category && category !== 'all') {
            products = products.filter(p => p.category === category);
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
    
    let subtotal = 0;
    let msg = '🛍️ <b>سلة المشتريات</b>\n━━━━━━━━━━━━━━━\n\n';
    const buttons = [];
    
    for (let i = 0; i < cart.length; i++) {
        const item = cart[i];
        subtotal += item.total;
        msg += `${i+1}. ${item.display}\n`;
        buttons.push([
            Markup.button.callback('➕', `inc_${i}`),
            Markup.button.callback(`${item.quantity}`, 'noop'),
            Markup.button.callback('➖', `dec_${i}`),
            Markup.button.callback('❌', `rem_${i}`)
        ]);
    }
    
    const delivery = subtotal >= FREE_DELIVERY_AMOUNT ? 0 : DELIVERY_FEE;
    const total = subtotal + delivery;
    const session = sessions.get(userId);
    const tier = session ? session.getTier() : { discount: 0 };
    const tierDiscount = Math.floor(subtotal * tier.discount / 100);
    const finalTotal = total - tierDiscount;
    
    msg += `\n━━━━━━━━━━━━━━━\n`;
    msg += `💰 <b>المجموع:</b> ${subtotal} ج\n`;
    if (delivery > 0) msg += `🚚 <b>التوصيل:</b> ${delivery} ج\n`;
    if (tierDiscount > 0) msg += `🎖️ <b>خصم ${tier.discount}%:</b> -${tierDiscount} ج\n`;
    msg += `━━━━━━━━━━━━━━━\n`;
    msg += `💎 <b>الإجمالي:</b> ${finalTotal} ج\n`;
    
    if (subtotal >= FREE_DELIVERY_AMOUNT) {
        msg += `\n🎉 <b>التوصيل مجاني!</b>`;
    }
    
    buttons.push([Markup.button.callback('➕ إضافة منتجات', 'browse_products')]);
    buttons.push([Markup.button.callback('🎟️ كود خصم', 'apply_coupon')]);
    buttons.push([Markup.button.callback('✅ تأكيد الطلب', 'checkout')]);
    buttons.push([Markup.button.callback('🗑️ تفريغ السلة', 'clear_cart')]);
    buttons.push([Markup.button.callback('🏠 الرئيسية', 'main_menu')]);
    
    await ctx.reply(msg, { parse_mode: 'HTML', ...Markup.inlineKeyboard(buttons) });
}

async function showOrders(ctx) {
    const userId = ctx.from.id;
    const orders = await getUserOrders(userId, MAX_ORDERS_DISPLAY);
    
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
    
    for (let i = 0; i < orders.length; i++) {
        const order = orders[i];
        let icon = order.status === 'تم التسليم' ? '✅' : order.status === 'في الطريق' ? '🚚' : order.status === 'ملغي' ? '❌' : '🟡';
        let statusText = order.status === 'جاري التنفيذ' ? 'جاري التحضير' : order.status;
        
        msg += `${i+1}. ${icon} <b>${order.date}</b>\n`;
        msg += `   🏷️ ${order.orderNumber}\n`;
        msg += `   💰 ${order.netPrice} ج\n`;
        msg += `   📍 ${statusText}\n`;
        msg += `━━━━━━━━━━━━━━━\n\n`;
        buttons.push([Markup.button.callback(`🔍 تتبع ${order.orderNumber.slice(-8)}`, `track_${order.orderNumber}`)]);
    }
    
    buttons.push([Markup.button.callback('📊 إحصائياتي', 'my_stats')]);
    buttons.push([Markup.button.callback('🏠 الرئيسية', 'main_menu')]);
    
    await ctx.reply(msg, { parse_mode: 'HTML', ...Markup.inlineKeyboard(buttons) });
}

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
        msg += `✨ ${offer.text}\n📅 ${offer.date}\n`;
        if (offer.code) msg += `🏷️ <b>كود:</b> ${offer.code}\n`;
        if (offer.discount > 0) msg += `💰 <b>خصم:</b> ${offer.discount}%\n`;
        msg += `━━━━━━━━━━━━━━━\n`;
    }
    
    await ctx.reply(msg, {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard([
            [Markup.button.callback('🛒 المنتجات', 'browse_products')],
            [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
        ])
    });
}

async function showStats(ctx) {
    const userId = ctx.from.id;
    const orders = await getUserOrders(userId, 100);
    const session = sessions.get(userId) || new UserSession(userId);
    const tier = session.getTier();
    
    const total = orders.reduce((sum, o) => sum + (parseFloat(o.netPrice) || 0), 0);
    const pending = orders.filter(o => o.status === 'جاري التنفيذ').length;
    const completed = orders.filter(o => o.status === 'تم التسليم').length;
    
    const msg = `📊 <b>إحصائياتي</b>\n━━━━━━━━━━━━━━━\n\n` +
        `👤 <b>الاسم:</b> ${session.name || 'غير مسجل'}\n` +
        `🏆 <b>المستوى:</b> ${tier.name} (خصم ${tier.discount}%)\n` +
        `📦 <b>عدد الطلبات:</b> ${orders.length}\n` +
        `✅ <b>تم التسليم:</b> ${completed}\n` +
        `🔄 <b>قيد التنفيذ:</b> ${pending}\n` +
        `💰 <b>إجمالي المشتريات:</b> ${total} ج\n` +
        `⭐ <b>نقاط الولاء:</b> ${session.loyaltyPoints}\n` +
        `🎯 <b>للوصول للمستوى التالي:</b> ${tier.name === '💎 بلاتينيوم' ? 'أنت في أعلى مستوى!' : `أنفق ${500 - total > 0 ? 500 - total : 0} ج أخرى`}`;
    
    await ctx.reply(msg, {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard([
            [Markup.button.callback('📋 طلباتي', 'my_orders')],
            [Markup.button.callback('⭐ تقييم الخدمة', 'feedback')],
            [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
        ])
    });
}

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
    
    let icon = '🟡', statusText = '', progress = 0;
    if (order.status === 'جاري التنفيذ') {
        icon = '🟡'; statusText = '✅ جاري تجهيز طلبك'; progress = 25;
    } else if (order.status === 'تم التجهيز') {
        icon = '📦'; statusText = 'طلبك جاهز للتوصيل'; progress = 50;
    } else if (order.status === 'في الطريق') {
        icon = '🚚'; statusText = 'مندوب التوصيل في طريقه إليك'; progress = 75;
    } else if (order.status === 'تم التسليم') {
        icon = '✅'; statusText = 'تم توصيل طلبك بنجاح'; progress = 100;
    } else if (order.status === 'ملغي') {
        icon = '❌'; statusText = 'تم إلغاء الطلب'; progress = 0;
    }
    
    const progressBar = '█'.repeat(Math.floor(progress / 10)) + '░'.repeat(10 - Math.floor(progress / 10));
    
    let msg = `📦 <b>تتبع الطلب</b>\n━━━━━━━━━━━━━━━\n\n`;
    msg += `🏷️ <b>رقم الطلب:</b> ${order.orderNumber}\n`;
    msg += `${icon} <b>الحالة:</b> ${order.status}\n`;
    msg += `📝 ${statusText}\n`;
    msg += `📊 التقدم: ${progressBar} ${progress}%\n\n`;
    msg += `👤 <b>العميل:</b> ${order.name}\n`;
    msg += `💰 <b>الإجمالي:</b> ${order.netPrice} ج\n`;
    if (order.discount > 0) msg += `🎖️ <b>الخصم:</b> ${order.discount} ج\n`;
    msg += `📍 <b>العنوان:</b> ${order.address}\n`;
    msg += `📅 <b>التاريخ:</b> ${order.date}\n`;
    msg += `━━━━━━━━━━━━━━━\n\n`;
    msg += `📦 <b>المنتجات:</b>\n${order.text}`;
    
    const buttons = [[Markup.button.callback('🏠 الرئيسية', 'main_menu')]];
    if (order.status === 'تم التسليم') {
        buttons.unshift([Markup.button.callback('⭐ تقييم الطلب', `rate_order_${order.orderNumber}`)]);
    }
    
    await ctx.reply(msg, { parse_mode: 'HTML', ...Markup.inlineKeyboard(buttons) });
}

async function startCheckout(ctx) {
    const userId = ctx.from.id;
    let cart = carts.get(userId) || [];
    
    if (cart.length === 0) {
        await ctx.reply('🛍️ سلتك فارغة');
        return;
    }
    
    if (cart.length > MAX_CART_ITEMS) {
        await ctx.reply(`❌ لا يمكنك طلب أكثر من ${MAX_CART_ITEMS} منتج في المرة الواحدة`);
        return;
    }
    
    const session = sessions.get(userId) || new UserSession(userId);
    session.status = 'ordering_name';
    sessions.set(userId, session);
    
    let subtotal = cart.reduce((sum, item) => sum + item.total, 0);
    const delivery = subtotal >= FREE_DELIVERY_AMOUNT ? 0 : DELIVERY_FEE;
    const tier = session.getTier();
    const discount = Math.floor(subtotal * tier.discount / 100);
    const total = subtotal + delivery - discount;
    
    const cartItems = cart.map((item, i) => `${i+1}. ${item.display}`).join('\n');
    
    await ctx.reply(`🛍️ <b>تأكيد الطلب</b>\n━━━━━━━━━━━━━━━\n\n${cartItems}\n\n` +
        `💰 المجموع: ${subtotal} ج\n` +
        (delivery > 0 ? `🚚 التوصيل: ${delivery} ج\n` : `🎉 توصيل مجاني!\n`) +
        (discount > 0 ? `🎖️ خصم ${tier.discount}%: -${discount} ج\n` : '') +
        `━━━━━━━━━━━━━━━\n` +
        `💎 <b>الإجمالي:</b> ${total} ج\n\n` +
        `📝 أرسل اسمك الكامل:`, { parse_mode: 'HTML' });
}

async function applyCoupon(ctx, code) {
    const offers = await getOffers();
    const offer = offers.find(o => o.code && o.code.toLowerCase() === code.toLowerCase());
    if (!offer) {
        await ctx.reply('❌ كود الخصم غير صالح');
        return false;
    }
    const session = sessions.get(ctx.from.id);
    if (session) {
        session.tempDiscount = offer.discount;
        sessions.set(ctx.from.id, session);
        await ctx.reply(`✅ تم تطبيق خصم ${offer.discount}% باستخدام الكود: ${code}`);
        return true;
    }
    return false;
}

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

// ============ القوائم الرئيسية ============
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

const adminKeyboard = Markup.inlineKeyboard([
    [Markup.button.callback('📊 إحصائيات النظام', 'admin_stats')],
    [Markup.button.callback('📢 بث رسالة', 'admin_broadcast')],
    [Markup.button.callback('🔄 تحديث الكاش', 'admin_refresh')]
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
    
    const welcomeMsg = `🌟 <b>مرحباً بك في بوب مارت!</b> 🌟\n━━━━━━━━━━━━━━━\n\n` +
        `👤 <b>الاسم:</b> ${ctx.from.first_name}\n` +
        `🏆 <b>المستوى:</b> ${session.getTier().name}\n\n` +
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
    await ctx.reply('🔍 <b>تتبع الطلب</b>\n\n✏️ أرسل رقم الطلب:', {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard([[Markup.button.callback('🏠 الرئيسية', 'main_menu')]])
    });
    const session = sessions.get(ctx.from.id) || new UserSession(ctx.from.id);
    session.status = 'tracking';
    sessions.set(ctx.from.id, session);
});

bot.action('feedback', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await showFeedbackButtons(ctx);
});

bot.action('support', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await ctx.reply('📞 <b>الدعم الفني</b>\n━━━━━━━━━━━━━━━\n\n' +
        '📱 واتساب: <code>01020063819</code>\n' +
        '✈️ تليجرام: @Muhamedhosny\n' +
        '📧 إيميل: support@bobmart.com\n\n' +
        '⏰ أوقات العمل: 9 ص - 9 م\n━━━━━━━━━━━━━━━\n\n' +
        '⚡ <b>للشكاوي والاقتراحات:</b> اضغط على زر التقييم', {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard([[Markup.button.callback('⭐ تقييم الخدمة', 'feedback')]])
    });
});

bot.action('search_products', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await ctx.reply('🔍 اكتب اسم المنتج:', {
        ...Markup.inlineKeyboard([[Markup.button.callback('🏠 الرئيسية', 'main_menu')]])
    });
    const session = sessions.get(ctx.from.id) || new UserSession(ctx.from.id);
    session.status = 'searching';
    sessions.set(ctx.from.id, session);
});

bot.action('apply_coupon', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await ctx.reply('🎟️ <b>إدخال كود الخصم</b>\n\n✏️ أرسل الكود:', {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard([[Markup.button.callback('🏠 الرئيسية', 'main_menu')]])
    });
    const session = sessions.get(ctx.from.id) || new UserSession(ctx.from.id);
    session.status = 'applying_coupon';
    sessions.set(ctx.from.id, session);
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
    
    if (session && session.status === 'editing') {
        const existing = session.editCart.find(item => item.name === productName);
        if (existing) existing.quantity++;
        else session.editCart.push({ name: productName, price: productPrice, quantity: 1 });
        sessions.set(userId, session);
        await ctx.reply(`✅ تم إضافة ${productName} للتعديل\n📦 المنتجات المضافة: ${session.editCart.map(i => `${i.quantity}×${i.name}`).join(', ')}`);
    } else {
        let cart = carts.get(userId) || [];
        if (cart.length >= MAX_CART_ITEMS) {
            await ctx.reply(`❌ لا يمكنك إضافة أكثر من ${MAX_CART_ITEMS} منتج`);
            return;
        }
        const existing = cart.find(item => item.name === productName);
        if (existing) existing.quantity++;
        else cart.push(new CartItem(productName, productPrice, 1));
        carts.set(userId, cart);
        await ctx.reply(`✅ تم إضافة ${productName} إلى السلة\n🛍️ عدد المنتجات في السلة: ${cart.length}`);
    }
});

bot.action('noop', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
});

// أزرار التقييم
for (let i = 1; i <= 5; i++) {
    bot.action(new RegExp(`rate_${i}(?:_(.*))?`), async (ctx) => {
        const rating = i;
        const orderNumber = ctx.match[1] || '';
        const userId = ctx.from.id;
        const session = sessions.get(userId) || new UserSession(userId);
        
        await ctx.answerCbQuery(`✅ شكراً لتقييمك ${rating} نجوم!`).catch(() => {});
        await saveFeedback(userId, session.name || ctx.from.first_name, rating, '', orderNumber);
        
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
    await ctx.reply('📝 <b>أضف ملاحظاتك</b>\n\n✏️ اكتب ملاحظاتك أو اقتراحاتك:');
});

bot.action(/rate_order_(.+)/, async (ctx) => {
    const orderNumber = ctx.match[1];
    await ctx.answerCbQuery().catch(() => {});
    const session = sessions.get(ctx.from.id) || new UserSession(ctx.from.id);
    session.pendingRating = orderNumber;
    sessions.set(ctx.from.id, session);
    await showFeedbackButtons(ctx, orderNumber);
});

// أوامر الأدمن
bot.action('admin_stats', async (ctx) => {
    if (ctx.from.id.toString() !== ADMIN_ID) return;
    const stats = await getSystemStats();
    await ctx.reply(`📊 <b>إحصائيات النظام</b>\n━━━━━━━━━━━━━━━\n\n` +
        `📦 إجمالي الطلبات: ${stats.totalOrders}\n` +
        `💰 إجمالي الإيرادات: ${stats.totalRevenue} ج\n` +
        `🔄 طلبات قيد التنفيذ: ${stats.pendingOrders}\n` +
        `⭐ متوسط التقييم: ${stats.avgRating}/5\n` +
        `👥 عدد الجلسات النشطة: ${sessions.size}\n` +
        `🛍️ عدد السلات النشطة: ${carts.size}`, { parse_mode: 'HTML' });
});

bot.action('admin_broadcast', async (ctx) => {
    if (ctx.from.id.toString() !== ADMIN_ID) return;
    const session = sessions.get(ctx.from.id);
    session.status = 'broadcast';
    sessions.set(ctx.from.id, session);
    await ctx.reply('📢 أرسل الرسالة التي تريد بثها لجميع العملاء:');
});

bot.action('admin_refresh', async (ctx) => {
    if (ctx.from.id.toString() !== ADMIN_ID) return;
    await getProducts(true);
    await getOffers(true);
    await ctx.reply('✅ تم تحديث الكاش بنجاح');
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
    
    // أوامر الأدمن
    if (text === '/admin' && ctx.from.id.toString() === ADMIN_ID) {
        await ctx.reply('🔐 لوحة تحكم الأدمن', { ...adminKeyboard });
        return;
    }
    
    // ملاحظات التقييم
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
    
    // بث رسالة للأدمن
    if (session.status === 'broadcast' && ctx.from.id.toString() === ADMIN_ID) {
        const customers = await getAllCustomers();
        let sent = 0;
        for (const customer of customers) {
            try {
                await bot.telegram.sendMessage(customer.id, `📢 <b>إعلان من بوب مارت</b>\n━━━━━━━━━━━━━━━\n\n${text}`, { parse_mode: 'HTML' });
                sent++;
                await new Promise(r => setTimeout(r, 50));
            } catch (err) { console.error(err); }
        }
        session.status = 'main';
        sessions.set(userId, session);
        await ctx.reply(`✅ تم إرسال الرسالة لـ ${sent} عميل`);
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
    
    // تطبيق كود خصم
    if (session.status === 'applying_coupon') {
        session.status = 'main';
        sessions.set(userId, session);
        await applyCoupon(ctx, text);
        await showCart(ctx);
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
        let subtotal = cart.reduce((sum, item) => sum + item.total, 0);
        const delivery = subtotal >= FREE_DELIVERY_AMOUNT ? 0 : DELIVERY_FEE;
        const tier = session.getTier();
        const discount = Math.floor(subtotal * tier.discount / 100);
        const total = subtotal + delivery - discount;
        
        const saved = await saveOrder(userId, session.name, session.address, session.phone, cart, subtotal, discount, total);
        
        if (saved.success) {
            carts.set(userId, []);
            session.status = 'main';
            session.addLoyaltyPoints(total);
            session.totalSpent += total;
            sessions.set(userId, session);
            
            const successMessage = 
`✅ <b>تم تسجيل طلبك بنجاح!</b>
━━━━━━━━━━━━━━━

🏷️ <b>رقم الطلب:</b> <code>${saved.orderNumber}</code>

👤 <b>الاسم:</b> ${session.name}
💰 <b>الإجمالي:</b> ${total} ج
${discount > 0 ? `🎖️ <b>الخصم:</b> ${discount} ج\n` : ''}
📍 <b>العنوان:</b> ${session.address}
📞 <b>الهاتف:</b> ${session.phone}

✨ <b>حصلت على ${Math.floor(total / 10)} نقطة ولاء</b>
🏆 <b>مستواك الحالي:</b> ${tier.name}

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

// ============ فحص العروض التلقائي ============
setInterval(async () => {
    console.log('🔍 جاري فحص العروض الجديدة...');
    await broadcastOffersToAllCustomers();
}, OFFERS_CHECK_INTERVAL);

setTimeout(async () => {
    console.log('🔍 فحص أولي للعروض...');
    await broadcastOffersToAllCustomers();
    await getProducts(true);
    console.log('✅ تم تحميل البوت بنجاح');
}, 30000);

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

// تشغيل محلي للتطوير
if (process.env.NODE_ENV !== 'production') {
    bot.launch();
    console.log('🤖 Bob Mart Bot is running locally...');
}

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
