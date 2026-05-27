import { Telegraf, Markup } from 'telegraf';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

const bot = new Telegraf(process.env.BOT_TOKEN);

// ============ الإعدادات ============
const ADMIN_ID = process.env.ADMIN_CHAT_ID || 'Muhamedhosny';
const ITEMS_PER_PAGE = 6;
const CACHE_TTL = 300000;
const CART_EXPIRY_HOURS = 48;
const OFFERS_CHECK_INTERVAL = 6 * 60 * 60 * 1000;

// ============ قاعدة البيانات المؤقتة ============
const sessions = new Map();
const carts = new Map();
const userCooldowns = new Map();
const offeredCustomers = new Map();
const usedCoupons = new Map();
let lastOffersHash = '';
let isBroadcasting = false;
let currentPage = 0;
let currentProducts = [];
let currentSearch = '';
let currentCategory = '';

// ============ كاش البيانات ============
let productsCache = null;
let productsCacheTime = 0;
let offersCache = null;
let offersCacheTime = 0;
let couponsCache = null;
let couponsCacheTime = 0;
let categoriesCache = null;
let categoriesCacheTime = 0;

// ============ دالة إرسال إشعار للأدمن ============
async function notifyAdmin(message, parseMode = 'HTML') {
    try {
        await bot.telegram.sendMessage(ADMIN_ID, message, { parse_mode: parseMode });
        console.log('📨 تم إرسال إشعار للأدمن');
    } catch (error) {
        console.error('❌ فشل إرسال الإشعار للأدمن:', error.message);
    }
}

// ============ كلاس جلسة المستخدم ============
class UserSession {
    constructor(userId, name = '', username = '') {
        this.userId = userId;
        this.name = name;
        this.username = username;
        this.address = '';
        this.phone = '';
        this.status = 'main';
        this.lastActivity = Date.now();
        this.totalSpent = 0;
        this.loyaltyPoints = 0;
        this.lastOrderNumber = null;
        this.editingOrder = null;
        this.editCart = [];
        this.removeItems = [];
        this.tempOrderNumber = null;
        this.useSavedAddress = null;
        this.appliedCoupon = null;
        this.couponDiscount = 0;
        this.viewingEditCart = false;
        this.selectedCategory = '';
    }
    
    updateActivity() {
        this.lastActivity = Date.now();
    }
    
    getTier() {
        if (this.totalSpent >= 5000) return { name: '💎 بلاتينيوم', discount: 10, minSpent: 5000 };
        if (this.totalSpent >= 2000) return { name: '🌟 ذهبي', discount: 7, minSpent: 2000 };
        if (this.totalSpent >= 500) return { name: '⭐ فضي', discount: 5, minSpent: 500 };
        if (this.totalSpent >= 100) return { name: '🟤 برونزي', discount: 3, minSpent: 100 };
        return { name: '🟣 جديد', discount: 0, minSpent: 0 };
    }
    
    addLoyaltyPoints(amount) {
        const points = Math.floor(amount / 10);
        this.loyaltyPoints += points;
        return points;
    }
    
    subtractLoyaltyPoints(amount) {
        const points = Math.floor(amount / 10);
        this.loyaltyPoints = Math.max(0, this.loyaltyPoints - points);
        return points;
    }
    
    applyCoupon(couponCode, discountPercent) {
        this.appliedCoupon = couponCode;
        this.couponDiscount = discountPercent;
    }
    
    removeCoupon() {
        this.appliedCoupon = null;
        this.couponDiscount = 0;
    }
}

// ============ كلاس عنصر السلة ============
class CartItem {
    constructor(name, price, quantity = 1, category = '') {
        this.id = Date.now().toString();
        this.name = name;
        this.price = price;
        this.quantity = Math.min(quantity, 99);
        this.addedAt = Date.now();
        this.category = category;
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

// ============ دوال المنتجات (مع دعم الأقسام) ============
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
        const categoriesSet = new Set();
        
        for (const row of rows) {
            const name = row.get('المنتج');
            const category = row.get('القسم') || 'عام';
            const price = parseFloat(row.get('السعر'));
            const unit = row.get('الوحدة') || 'قطعة';
            const available = row.get('التوفر');
            
            let isAvailable = true;
            if (available === '0' || available === 'غير متوفر' || available === 'لا') {
                isAvailable = false;
            }
            
            if (name && price && !isNaN(price) && price > 0 && isAvailable) {
                products.push({ name, category, price, unit });
                categoriesSet.add(category);
            }
        }
        
        productsCache = products;
        productsCacheTime = now;
        categoriesCache = Array.from(categoriesSet).sort();
        categoriesCacheTime = now;
        
        console.log(`✅ تم تحميل ${products.length} منتج من ${categoriesCache.length} قسم`);
        return products;
    } catch (error) {
        console.error('❌ خطأ في المنتجات:', error);
        return productsCache || [];
    }
}

async function getCategories(forceRefresh = false) {
    const now = Date.now();
    if (!forceRefresh && categoriesCache && (now - categoriesCacheTime) < CACHE_TTL) {
        return categoriesCache;
    }
    await getProducts(forceRefresh);
    return categoriesCache || [];
}

async function getProductsByCategory(category) {
    const products = await getProducts();
    if (category === 'all' || !category) {
        return products;
    }
    return products.filter(p => p.category === category);
}

// ============ دوال العروض ============
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
            const price = row.get('السعر');
            const endDate = row.get('تاريخ الإنتهاء');
            
            if (active === 'نعم' && text) {
                offers.push({ 
                    text: text, 
                    price: price || 'غير محدد',
                    endDate: endDate || 'غير محدد'
                });
            }
        }
        offersCache = offers;
        offersCacheTime = now;
        console.log(`✅ تم تحميل ${offers.length} عرض`);
        return offers;
    } catch (error) {
        console.error('❌ خطأ في العروض:', error);
        return offersCache || [];
    }
}

// ============ دوال الكوبونات ============
async function getCoupons(forceRefresh = false) {
    const now = Date.now();
    if (!forceRefresh && couponsCache && (now - couponsCacheTime) < CACHE_TTL) {
        return couponsCache;
    }
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['Coupons'];
        if (!sheet) return [];
        const rows = await sheet.getRows();
        const coupons = [];
        
        for (const row of rows) {
            const code = row.get('Coupon');
            const discount = parseFloat(row.get('الخصم'));
            
            if (!code || !discount) continue;
            
            coupons.push({
                code: code.toUpperCase(),
                discount: discount
            });
        }
        
        couponsCache = coupons;
        couponsCacheTime = now;
        console.log(`✅ تم تحميل ${coupons.length} كوبون`);
        return coupons;
    } catch (error) {
        console.error('❌ خطأ في الكوبونات:', error);
        return couponsCache || [];
    }
}

async function validateCoupon(code, userId) {
    const coupons = await getCoupons();
    const coupon = coupons.find(c => c.code === code.toUpperCase());
    
    if (!coupon) {
        return { valid: false, message: '❌ الكوبون غير صالح' };
    }
    
    const userKey = `${userId}_${coupon.code}`;
    if (usedCoupons.has(userKey)) {
        return { valid: false, message: '❌ لقد استخدمت هذا الكوبون من قبل' };
    }
    
    return { 
        valid: true, 
        discount: coupon.discount, 
        code: coupon.code,
        message: `✅ تم تطبيق كود ${coupon.code} - خصم ${coupon.discount}%`
    };
}

async function markCouponAsUsed(userId, couponCode) {
    const userKey = `${userId}_${couponCode}`;
    usedCoupons.set(userKey, {
        usedAt: new Date().toISOString(),
        couponCode: couponCode
    });
}

// ============ دوال الطلبات ============
async function saveOrder(userId, customerName, address, phone, items, subtotal, finalTotal, loyaltyPoints) {
    try {
        const doc = await getDoc();
        let sheet = doc.sheetsByTitle['Orders'];
        if (!sheet) {
            sheet = await doc.addSheet({
                title: 'Orders',
                headerValues: ['ID', 'اسم العميل', 'الطلبات', 'الإجمالي', 'الإجمالي بعد الخصم', 'نقاط الولاء', 'العنوان', 'رقم الهاتف', 'التاريخ', 'عدد العناصر', 'رقم الطلب', 'الحالة']
            });
        }
        
        let calculatedTotal = 0;
        for (const item of items) {
            calculatedTotal += item.price * item.quantity;
        }
        
        const orderText = items.map(item => `${item.quantity} × ${item.name} (${item.price}ج) = ${item.price * item.quantity}ج`).join('\n• ');
        const timestamp = Date.now();
        const orderNumber = `ORD-${timestamp}-${userId.toString().slice(-4)}`;
        
        await sheet.addRow({
            'ID': userId.toString(),
            'اسم العميل': customerName,
            'الطلبات': `• ${orderText}`,
            'الإجمالي': calculatedTotal,
            'الإجمالي بعد الخصم': finalTotal,
            'نقاط الولاء': loyaltyPoints,
            'العنوان': address,
            'رقم الهاتف': phone,
            'التاريخ': new Date().toLocaleString('ar-EG'),
            'عدد العناصر': items.reduce((sum, i) => sum + i.quantity, 0),
            'رقم الطلب': orderNumber,
            'الحالة': 'جاري التنفيذ'
        });
        
        console.log(`✅ تم حفظ الطلب: ${orderNumber}`);
        return { success: true, orderNumber, finalTotal: finalTotal, subtotal: calculatedTotal };
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
                const subtotal = parseFloat(row.get('الإجمالي')) || 0;
                const finalTotal = parseFloat(row.get('الإجمالي بعد الخصم')) || subtotal;
                const discountAmount = subtotal - finalTotal;
                const loyaltyPoints = parseFloat(row.get('نقاط الولاء')) || 0;
                
                orders.push({
                    orderNumber: row.get('رقم الطلب') || '',
                    subtotal: subtotal,
                    finalTotal: finalTotal,
                    discountAmount: discountAmount > 0 ? discountAmount : 0,
                    loyaltyPoints: loyaltyPoints,
                    date: row.get('التاريخ'),
                    status: row.get('الحالة') || 'جاري التنفيذ',
                    itemsText: row.get('الطلبات'),
                    address: row.get('العنوان'),
                    phone: row.get('رقم الهاتف'),
                    itemCount: parseInt(row.get('عدد العناصر')) || 0
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
                const subtotal = parseFloat(row.get('الإجمالي')) || 0;
                const finalTotal = parseFloat(row.get('الإجمالي بعد الخصم')) || subtotal;
                const discountAmount = subtotal - finalTotal;
                const loyaltyPoints = parseFloat(row.get('نقاط الولاء')) || 0;
                
                return {
                    orderNumber: row.get('رقم الطلب'),
                    name: row.get('اسم العميل'),
                    itemsText: row.get('الطلبات'),
                    subtotal: subtotal,
                    finalTotal: finalTotal,
                    discountAmount: discountAmount > 0 ? discountAmount : 0,
                    loyaltyPoints: loyaltyPoints,
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

async function removeItemsFromOrder(orderNumber, itemsToRemove) {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['Orders'];
        if (!sheet) return false;
        const rows = await sheet.getRows();
        for (const row of rows) {
            if (row.get('رقم الطلب') === orderNumber) {
                let oldText = row.get('الطلبات');
                let oldTotal = parseFloat(row.get('الإجمالي')) || 0;
                let oldFinal = parseFloat(row.get('الإجمالي بعد الخصم')) || oldTotal;
                let oldCount = parseInt(row.get('عدد العناصر')) || 0;
                
                let removedTotal = 0;
                let removedCount = 0;
                let newText = oldText;
                
                for (const item of itemsToRemove) {
                    const pattern = new RegExp(`• \\d+ × ${item.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')} \\(${item.price}ج\\) = ${item.price * item.quantity}ج`, 'g');
                    newText = newText.replace(pattern, '');
                    removedTotal += item.price * item.quantity;
                    removedCount += item.quantity;
                }
                
                newText = newText.replace(/\n• \n/g, '\n').replace(/\n\n/g, '\n');
                if (newText.startsWith('• ')) {
                    newText = newText;
                } else if (newText.trim() === '') {
                    newText = '• لا توجد منتجات';
                }
                
                const newFinal = oldFinal - removedTotal;
                const newPoints = Math.floor(newFinal / 10);
                
                row.set('الطلبات', newText);
                row.set('الإجمالي', oldTotal - removedTotal);
                row.set('الإجمالي بعد الخصم', newFinal);
                row.set('نقاط الولاء', newPoints);
                row.set('عدد العناصر', oldCount - removedCount);
                await row.save();
                return true;
            }
        }
        return false;
    } catch (error) {
        console.error('❌ خطأ في حذف المنتجات:', error);
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
                const newItemsText = newItems.map(item => `${item.quantity} × ${item.name} (${item.price}ج) = ${item.price * item.quantity}ج`).join('\n• ');
                
                let newText = oldText;
                if (oldText === '• لا توجد منتجات') {
                    newText = `• ${newItemsText}`;
                } else {
                    newText = oldText + '\n• ' + newItemsText;
                }
                
                row.set('الطلبات', newText);
                
                const oldTotal = parseFloat(row.get('الإجمالي')) || 0;
                const newTotal = newItems.reduce((sum, item) => sum + (item.price * item.quantity), 0);
                row.set('الإجمالي', oldTotal + newTotal);
                
                const oldFinal = parseFloat(row.get('الإجمالي بعد الخصم')) || oldTotal;
                const newFinal = oldFinal + newTotal;
                row.set('الإجمالي بعد الخصم', newFinal);
                
                const newPoints = Math.floor(newFinal / 10);
                row.set('نقاط الولاء', newPoints);
                
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

async function getUserTotalSpent(userId) {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['Orders'];
        if (!sheet) return 0;
        const rows = await sheet.getRows();
        let total = 0;
        for (const row of rows) {
            if (row.get('ID') === userId.toString() && row.get('الحالة') !== 'ملغي') {
                let finalTotal = parseFloat(row.get('الإجمالي بعد الخصم')) || 0;
                if (finalTotal === 0) finalTotal = parseFloat(row.get('الإجمالي')) || 0;
                total += finalTotal;
            }
        }
        return total;
    } catch (error) {
        return 0;
    }
}

async function getUserLoyaltyPoints(userId) {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['Orders'];
        if (!sheet) return 0;
        const rows = await sheet.getRows();
        let total = 0;
        for (const row of rows) {
            if (row.get('ID') === userId.toString() && row.get('الحالة') !== 'ملغي') {
                total += parseFloat(row.get('نقاط الولاء')) || 0;
            }
        }
        return total;
    } catch (error) {
        return 0;
    }
}

async function getUserLastAddress(userId) {
    try {
        const orders = await getUserOrders(userId);
        const validOrders = orders.filter(o => o.status !== 'ملغي');
        if (validOrders.length > 0) {
            const lastOrder = validOrders[0];
            return { address: lastOrder.address, phone: lastOrder.phone, name: '' };
        }
        return null;
    } catch (error) {
        return null;
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
            if (id && name && !customersMap.has(id)) {
                customersMap.set(id, { id, name });
            }
        }
        return Array.from(customersMap.values());
    } catch (error) {
        console.error('❌ خطأ في جلب العملاء:', error);
        return [];
    }
}

// ============ دوال التقييم ============
async function saveFeedback(userId, userName, rating) {
    try {
        const doc = await getDoc();
        let sheet = doc.sheetsByTitle['Feedbacks'];
        if (!sheet) {
            sheet = await doc.addSheet({
                title: 'Feedbacks',
                headerValues: ['ID', 'الاسم', 'التقييم', 'التاريخ']
            });
        }
        await sheet.addRow({
            'ID': userId.toString(),
            'الاسم': userName,
            'التقييم': rating,
            'التاريخ': new Date().toLocaleString('ar-EG')
        });
        console.log(`✅ تم حفظ تقييم ${rating} نجوم من ${userName}`);
        return true;
    } catch (error) {
        console.error('❌ خطأ في حفظ التقييم:', error);
        return false;
    }
}

// ============ حساب الخصومات ============
async function calculateDiscounts(userId, subtotal, appliedCouponCode = null) {
    const totalSpent = await getUserTotalSpent(userId);
    let tierDiscountPercent = 0;
    if (totalSpent >= 5000) tierDiscountPercent = 10;
    else if (totalSpent >= 2000) tierDiscountPercent = 7;
    else if (totalSpent >= 500) tierDiscountPercent = 5;
    else if (totalSpent >= 100) tierDiscountPercent = 3;
    
    const tierDiscountAmount = Math.floor(subtotal * tierDiscountPercent / 100);
    
    let couponDiscountAmount = 0;
    let couponValid = false;
    let couponMessage = '';
    
    if (appliedCouponCode) {
        const couponCheck = await validateCoupon(appliedCouponCode, userId);
        if (couponCheck.valid) {
            couponDiscountAmount = Math.floor(subtotal * couponCheck.discount / 100);
            couponValid = true;
            couponMessage = couponCheck.message;
        } else {
            couponMessage = couponCheck.message;
        }
    }
    
    let finalDiscountAmount = tierDiscountAmount;
    let usedCoupon = null;
    let usedTier = true;
    
    if (couponValid && couponDiscountAmount > tierDiscountAmount) {
        finalDiscountAmount = couponDiscountAmount;
        usedCoupon = appliedCouponCode;
        usedTier = false;
    }
    
    let finalTotal = subtotal - finalDiscountAmount;
    if (finalTotal <= 0 && subtotal > 0) {
        finalTotal = subtotal;
        finalDiscountAmount = 0;
        usedCoupon = null;
    }
    
    return {
        subtotal: subtotal,
        tierDiscount: { percent: tierDiscountPercent, amount: tierDiscountAmount, applied: usedTier },
        couponDiscount: { percent: couponValid ? Math.round(couponDiscountAmount * 100 / subtotal) : 0, amount: couponDiscountAmount, valid: couponValid, message: couponMessage, applied: !usedTier && couponValid },
        finalTotal: finalTotal,
        usedCoupon: usedCoupon
    };
}

// ============ البث التلقائي للعروض ============
function getOffersHash(offers) {
    return JSON.stringify(offers.map(o => ({ text: o.text, price: o.price, endDate: o.endDate })));
}

async function broadcastOffersToAllCustomers() {
    if (isBroadcasting) return;
    
    try {
        isBroadcasting = true;
        const offers = await getOffers(true);
        const currentHash = getOffersHash(offers);
        
        if (currentHash === lastOffersHash || offers.length === 0) {
            console.log('📭 لا توجد عروض جديدة');
            return;
        }
        
        console.log('📢 تم اكتشاف عروض جديدة! جاري البث...');
        
        const customers = await getAllCustomers();
        let sentCount = 0;
        
        for (const customer of customers) {
            const customerKey = `${customer.id}_${currentHash}`;
            if (offeredCustomers.has(customerKey)) continue;
            
            try {
                let offersText = '';
                const offerButtons = [];
                for (let i = 0; i < Math.min(offers.length, 5); i++) {
                    const offer = offers[i];
                    offersText += `✨ ${offer.text}\n💰 السعر: ${offer.price}\n📅 ينتهي: ${offer.endDate}\n━━━━━━━━━━━━━━━\n`;
                    offerButtons.push([Markup.button.callback(`🛒 اطلب ${offer.text.substring(0, 20)}`, `order_offer_${i}`)]);
                }
                
                const message = `🎁 <b>عروض حصرية من سوبر ماركت الحَواج!</b> 🎁\n━━━━━━━━━━━━━━━\n\n${offersText}\n⬇️ <b>اطلب الآن واستفد من العروض!</b>`;
                
                await bot.telegram.sendMessage(customer.id, message, {
                    parse_mode: 'HTML',
                    ...Markup.inlineKeyboard([
                        ...offerButtons,
                        [Markup.button.callback('🛒 تصفح جميع المنتجات', 'browse_products')],
                        [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
                    ])
                });
                offeredCustomers.set(customerKey, true);
                sentCount++;
                await new Promise(r => setTimeout(r, 100));
            } catch (err) {
                console.error(`فشل الإرسال للعميل ${customer.id}:`, err.message);
            }
        }
        
        console.log(`✅ تم بث العروض لـ ${sentCount} عميل`);
        lastOffersHash = currentHash;
        
        setTimeout(() => {
            for (const key of offeredCustomers.keys()) {
                if (key.includes(currentHash)) offeredCustomers.delete(key);
            }
        }, 24 * 60 * 60 * 1000);
        
    } catch (error) {
        console.error('❌ خطأ في بث العروض:', error);
    } finally {
        isBroadcasting = false;
    }
}

// ============ عرض الأقسام ============
async function showCategories(ctx) {
    const categories = await getCategories();
    
    if (categories.length === 0) {
        await ctx.reply('⚠️ لا توجد أقسام متاحة حالياً', {
            ...Markup.inlineKeyboard([[Markup.button.callback('🏠 الرئيسية', 'main_menu')]])
        });
        return;
    }
    
    const buttons = [];
    for (const category of categories) {
        buttons.push([Markup.button.callback(`📁 ${category}`, `category_${category}`)]);
    }
    buttons.push([Markup.button.callback('📦 جميع المنتجات', 'browse_products')]);
    buttons.push([Markup.button.callback('🏠 الرئيسية', 'main_menu')]);
    
    await ctx.reply('📂 <b>اختر القسم الذي تريد تصفحه:</b>', {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard(buttons)
    });
}

// ============ عرض المنتجات (مع دعم الأقسام) ============
async function showProducts(ctx, page = 0, searchQuery = '', category = '') {
    try {
        let products;
        if (category && category !== 'all') {
            products = await getProductsByCategory(category);
        } else {
            products = await getProducts();
        }
        
        if (products.length === 0) {
            await ctx.reply('⚠️ لا توجد منتجات في هذا القسم', {
                ...Markup.inlineKeyboard([
                    [Markup.button.callback('📂 عرض الأقسام', 'show_categories')],
                    [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
                ])
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
        currentCategory = category;
        
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
        buttons.push([Markup.button.callback('📂 أقسام', 'show_categories')]);
        buttons.push([Markup.button.callback('🛍️ السلة', 'view_cart')]);
        buttons.push([Markup.button.callback('🏠 الرئيسية', 'main_menu')]);
        
        let msg = '🛒 <b>المنتجات</b>\n━━━━━━━━━━━━━━━\n';
        if (category) msg += `📂 القسم: ${category}\n`;
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
                [Markup.button.callback('📂 الأقسام', 'show_categories')],
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
        msg += `${i+1}. ${item.name}: ${item.quantity} × ${item.price} = ${item.total} ج\n`;
        buttons.push([
            Markup.button.callback('➕', `inc_${i}`),
            Markup.button.callback(`${item.quantity}`, 'noop'),
            Markup.button.callback('➖', `dec_${i}`),
            Markup.button.callback('❌', `rem_${i}`)
        ]);
    }
    
    const session = sessions.get(userId);
    const discounts = await calculateDiscounts(userId, subtotal, session?.appliedCoupon);
    
    msg += `\n━━━━━━━━━━━━━━━\n`;
    msg += `💰 <b>المجموع:</b> ${discounts.subtotal} ج\n`;
    if (discounts.tierDiscount.amount > 0 && discounts.tierDiscount.applied) {
        msg += `🏆 <b>خصم المستوى (${discounts.tierDiscount.percent}%):</b> -${discounts.tierDiscount.amount} ج\n`;
    }
    if (discounts.couponDiscount.valid && discounts.couponDiscount.applied && discounts.couponDiscount.amount > 0) {
        msg += `🎟️ <b>خصم الكوبون (${discounts.couponDiscount.percent}%):</b> -${discounts.couponDiscount.amount} ج\n`;
    }
    msg += `━━━━━━━━━━━━━━━\n`;
    msg += `💎 <b>الإجمالي بعد الخصم:</b> ${discounts.finalTotal} ج`;
    
    const points = Math.floor(discounts.finalTotal / 10);
    msg += `\n⭐ <b>نقاط الولاء المتوقعة:</b> ${points} نقطة`;
    
    buttons.push([Markup.button.callback('➕ إضافة منتجات', 'browse_products')]);
    buttons.push([Markup.button.callback('📂 أقسام', 'show_categories')]);
    buttons.push([Markup.button.callback('🎟️ إدخال كوبون', 'enter_coupon')]);
    buttons.push([Markup.button.callback('✅ تأكيد الطلب', 'checkout')]);
    buttons.push([Markup.button.callback('🗑️ تفريغ السلة', 'clear_cart')]);
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
                [Markup.button.callback('📂 الأقسام', 'show_categories')],
                [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
            ])
        });
        return;
    }
    
    const buttons = [];
    let msg = '🎁 <b>عروض سوبر ماركت الحَواج</b>\n━━━━━━━━━━━━━━━\n\n';
    
    for (let i = 0; i < offers.length; i++) {
        const offer = offers[i];
        msg += `✨ <b>${offer.text}</b>\n`;
        msg += `💰 السعر: ${offer.price}\n`;
        msg += `📅 ينتهي: ${offer.endDate}\n`;
        msg += `━━━━━━━━━━━━━━━\n`;
        buttons.push([Markup.button.callback(`🛒 اطلب ${offer.text.substring(0, 25)}`, `order_offer_${i}`)]);
    }
    
    buttons.push([Markup.button.callback('🛒 تصفح جميع المنتجات', 'browse_products')]);
    buttons.push([Markup.button.callback('📂 الأقسام', 'show_categories')]);
    buttons.push([Markup.button.callback('🏠 الرئيسية', 'main_menu')]);
    
    await ctx.reply(msg, {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard(buttons)
    });
}

// ============ عرض الكوبونات ============
async function showAvailableCoupons(ctx) {
    const coupons = await getCoupons();
    
    if (coupons.length === 0) {
        await ctx.reply('🎟️ لا توجد كوبونات خصم متاحة حالياً', {
            ...Markup.inlineKeyboard([[Markup.button.callback('🏠 الرئيسية', 'main_menu')]])
        });
        return;
    }
    
    let msg = '🎟️ <b>كوبونات الخصم المتاحة</b>\n━━━━━━━━━━━━━━━\n\n';
    for (const coupon of coupons) {
        msg += `🏷️ <b>الكود:</b> <code>${coupon.code}</code>\n`;
        msg += `💰 <b>الخصم:</b> ${coupon.discount}%\n`;
        msg += `━━━━━━━━━━━━━━━\n`;
    }
    msg += `\n💡 <b>كيف تستخدم الكوبون؟</b>\n`;
    msg += `1. أضف منتجاتك للسلة\n`;
    msg += `2. اضغط "إدخال كوبون"\n`;
    msg += `3. اكتب الكوبون مثل: <code>${coupons[0].code}</code>`;
    
    await ctx.reply(msg, {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard([
            [Markup.button.callback('🛒 التسوق الآن', 'browse_products')],
            [Markup.button.callback('📂 الأقسام', 'show_categories')],
            [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
        ])
    });
}

// ============ إدخال كوبون ============
async function enterCoupon(ctx) {
    await ctx.reply('🎟️ <b>إدخال كود الخصم</b>\n\n✏️ أرسل الكود (مثال: WELCOME10):\n\n📌 لعرض الكوبونات المتاحة اضغط "عرض الكوبونات"', {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard([
            [Markup.button.callback('🎁 عرض الكوبونات', 'show_coupons')],
            [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
        ])
    });
    
    const session = sessions.get(ctx.from.id) || new UserSession(ctx.from.id);
    session.status = 'entering_coupon';
    sessions.set(ctx.from.id, session);
}

// ============ معالج طلب عرض ============
async function handleOrderOffer(ctx, offerIndex) {
    const offers = await getOffers();
    const offer = offers[parseInt(offerIndex)];
    
    if (!offer) {
        await ctx.reply('❌ العرض غير موجود');
        return;
    }
    
    let productName = offer.text;
    let productPrice = parseFloat(offer.price);
    
    if (isNaN(productPrice)) {
        productPrice = 0;
    }
    
    const userId = ctx.from.id;
    let cart = carts.get(userId) || [];
    
    const existing = cart.find(item => item.name === productName);
    if (existing) {
        existing.quantity++;
    } else {
        cart.push(new CartItem(productName, productPrice, 1));
    }
    carts.set(userId, cart);
    
    await ctx.answerCbQuery(`✅ تم إضافة ${productName} إلى السلة`).catch(() => {});
    
    await ctx.reply(`✅ تم إضافة <b>${productName}</b> إلى السلة\n💰 السعر: ${productPrice} ج`, {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard([
            [Markup.button.callback('🛍️ عرض السلة', 'view_cart')],
            [Markup.button.callback('🛒 متابعة التسوق', 'browse_products')],
            [Markup.button.callback('📂 الأقسام', 'show_categories')],
            [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
        ])
    });
}

// ============ عرض الطلبات ============
async function showOrders(ctx) {
    const userId = ctx.from.id;
    const orders = await getUserOrders(userId);
    
    if (orders.length === 0) {
        await ctx.reply('📭 لا يوجد طلبات سابقة', {
            ...Markup.inlineKeyboard([
                [Markup.button.callback('🛒 المنتجات', 'browse_products')],
                [Markup.button.callback('📂 الأقسام', 'show_categories')],
                [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
            ])
        });
        return;
    }
    
    const buttons = [];
    
    for (let i = 0; i < orders.length; i++) {
        const order = orders[i];
        let icon = order.status === 'تم التسليم' ? '✅' : order.status === 'في الطريق' ? '🚚' : order.status === 'ملغي' ? '❌' : '🟡';
        
        let msg = `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
        msg += `📋 <b>الطلب #${i + 1}</b>\n`;
        msg += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
        msg += `${icon} <b>الحالة:</b> ${order.status}\n`;
        msg += `🏷️ <b>رقم الطلب:</b> <code>${order.orderNumber}</code>\n`;
        msg += `📅 <b>التاريخ:</b> ${order.date}\n`;
        msg += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
        msg += `📦 <b>المنتجات:</b>\n${order.itemsText || 'لا توجد منتجات'}\n`;
        msg += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
        msg += `💰 <b>الإجمالي:</b> ${order.subtotal} ج\n`;
        if (order.discountAmount > 0) {
            msg += `🎁 <b>الخصم:</b> -${order.discountAmount} ج\n`;
        }
        msg += `💎 <b>الإجمالي بعد الخصم:</b> ${order.finalTotal} ج\n`;
        msg += `⭐ <b>نقاط الولاء المكتسبة:</b> ${order.loyaltyPoints} نقطة\n`;
        msg += `📍 <b>العنوان:</b> ${order.address}\n`;
        msg += `📞 <b>الهاتف:</b> ${order.phone}\n`;
        msg += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
        
        await ctx.reply(msg, { parse_mode: 'HTML' });
        
        buttons.push([Markup.button.callback(`🔍 تتبع الطلب ${i + 1}`, `track_${order.orderNumber}`)]);
    }
    
    buttons.push([Markup.button.callback('📊 إحصائياتي', 'my_stats')]);
    buttons.push([Markup.button.callback('🏠 الرئيسية', 'main_menu')]);
    
    await ctx.reply('📌 <b>اختر الإجراء المناسب:</b>', {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard(buttons)
    });
}

// ============ تتبع الطلب ============
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
    msg += `🏷️ <b>رقم الطلب:</b> <code>${order.orderNumber}</code>\n`;
    msg += `${icon} <b>الحالة:</b> ${order.status}\n`;
    msg += `📝 ${statusText}\n\n`;
    msg += `👤 <b>العميل:</b> ${order.name}\n`;
    msg += `💰 <b>الإجمالي:</b> ${order.subtotal} ج\n`;
    if (order.discountAmount > 0) {
        msg += `🎁 <b>الخصم:</b> -${order.discountAmount} ج\n`;
    }
    msg += `💎 <b>الإجمالي المدفوع:</b> ${order.finalTotal} ج\n`;
    msg += `⭐ <b>نقاط الولاء المكتسبة:</b> ${order.loyaltyPoints} نقطة\n`;
    msg += `📍 <b>العنوان:</b> ${order.address}\n`;
    msg += `📞 <b>الهاتف:</b> ${order.phone}\n`;
    msg += `📅 <b>التاريخ:</b> ${order.date}\n`;
    msg += `━━━━━━━━━━━━━━━\n\n`;
    msg += `📦 <b>المنتجات:</b>\n${order.itemsText}`;
    
    const buttons = [];
    
    if (order.canEdit) {
        buttons.push([Markup.button.callback('✏️ تعديل الطلب (إضافة منتجات)', `edit_order_${order.orderNumber}`)]);
        buttons.push([Markup.button.callback('🗑️ حذف منتجات من الطلب', `delete_items_${order.orderNumber}`)]);
        buttons.push([Markup.button.callback('❌ إلغاء الطلب كاملاً', `cancel_order_${order.orderNumber}`)]);
    }
    
    if (order.status === 'تم التسليم') {
        buttons.push([Markup.button.callback('⭐ تقييم الطلب', `rate_order_${order.orderNumber}`)]);
    }
    
    buttons.push([Markup.button.callback('📋 كل طلباتي', 'my_orders')]);
    buttons.push([Markup.button.callback('🏠 الرئيسية', 'main_menu')]);
    
    await ctx.reply(msg, { parse_mode: 'HTML', ...Markup.inlineKeyboard(buttons) });
}

// ============ عرض منتجات الطلب للحذف ============
async function showOrderItemsForDeletion(ctx, orderNumber) {
    const order = await getOrderByNumber(orderNumber);
    
    if (!order || !order.canEdit) {
        await ctx.reply('❌ لا يمكن تعديل هذا الطلب حالياً');
        return;
    }
    
    const items = [];
    const lines = order.itemsText.split('\n');
    for (const line of lines) {
        const match = line.match(/(\d+) × (.+?) \((\d+)ج\) = (\d+)ج/);
        if (match) {
            items.push({
                name: match[2],
                price: parseInt(match[3]),
                quantity: parseInt(match[1]),
                total: parseInt(match[4])
            });
        }
    }
    
    if (items.length === 0) {
        await ctx.reply('❌ لا توجد منتجات لحذفها');
        return;
    }
    
    const session = sessions.get(ctx.from.id) || new UserSession(ctx.from.id);
    session.status = 'deleting_items';
    session.editingOrder = orderNumber;
    session.removeItems = [];
    sessions.set(ctx.from.id, session);
    
    let msg = `🗑️ <b>حذف منتجات من الطلب #${orderNumber}</b>\n━━━━━━━━━━━━━━━\n\n`;
    msg += `📦 <b>المنتجات الحالية:</b>\n\n`;
    
    const buttons = [];
    for (let i = 0; i < items.length; i++) {
        const item = items[i];
        msg += `${i+1}. ${item.name}: ${item.quantity} × ${item.price} = ${item.total} ج\n`;
        buttons.push([Markup.button.callback(`❌ حذف ${item.name}`, `remove_item_${i}`)]);
    }
    
    msg += `\n━━━━━━━━━━━━━━━\n`;
    msg += `⚠️ اضغط على المنتج الذي تريد حذفه\n`;
    msg += `يمكنك حذف عدة منتجات ثم الضغط على "تأكيد الحذف"`;
    
    buttons.push([Markup.button.callback('✅ تأكيد الحذف', 'confirm_remove_items')]);
    buttons.push([Markup.button.callback('❌ إلغاء', 'cancel_edit')]);
    buttons.push([Markup.button.callback('🏠 الرئيسية', 'main_menu')]);
    
    session.orderItems = items;
    sessions.set(ctx.from.id, session);
    
    await ctx.reply(msg, {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard(buttons)
    });
}

// ============ إضافة منتج للحذف ============
async function markItemForRemoval(ctx, itemIndex) {
    const userId = ctx.from.id;
    const session = sessions.get(userId);
    
    if (!session || session.status !== 'deleting_items') {
        await ctx.reply('❌ لا توجد عملية حذف نشطة');
        return;
    }
    
    const item = session.orderItems[parseInt(itemIndex)];
    if (!item) return;
    
    const alreadySelected = session.removeItems.find(i => i.name === item.name);
    if (alreadySelected) {
        await ctx.answerCbQuery(`⚠️ ${item.name} محدد بالفعل للحذف`).catch(() => {});
        return;
    }
    
    session.removeItems.push(item);
    sessions.set(userId, session);
    
    await ctx.answerCbQuery(`✅ تم تحديد ${item.name} للحذف`).catch(() => {});
    
    let msg = `🗑️ <b>حذف منتجات من الطلب #${session.editingOrder}</b>\n━━━━━━━━━━━━━━━\n\n`;
    msg += `📦 <b>المنتجات المحددة للحذف:</b>\n`;
    if (session.removeItems.length === 0) {
        msg += `❌ لم يتم تحديد أي منتج\n`;
    } else {
        for (const item of session.removeItems) {
            msg += `• ${item.name}: ${item.quantity} × ${item.price} = ${item.total} ج\n`;
        }
    }
    msg += `\n━━━━━━━━━━━━━━━\n`;
    msg += `✅ تم تحديد ${session.removeItems.length} منتج\n`;
    msg += `اضغط "تأكيد الحذف" لإتمام العملية`;
    
    const buttons = [];
    for (let i = 0; i < session.orderItems.length; i++) {
        const itm = session.orderItems[i];
        const isSelected = session.removeItems.find(r => r.name === itm.name);
        const prefix = isSelected ? '✅' : '❌';
        buttons.push([Markup.button.callback(`${prefix} ${itm.name}`, `remove_item_${i}`)]);
    }
    buttons.push([Markup.button.callback('✅ تأكيد الحذف', 'confirm_remove_items')]);
    buttons.push([Markup.button.callback('❌ إلغاء', 'cancel_edit')]);
    buttons.push([Markup.button.callback('🏠 الرئيسية', 'main_menu')]);
    
    await ctx.reply(msg, {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard(buttons)
    });
}

// ============ تأكيد حذف المنتجات ============
async function confirmRemoveItems(ctx) {
    const userId = ctx.from.id;
    const session = sessions.get(userId);
    
    if (!session || session.status !== 'deleting_items' || session.removeItems.length === 0) {
        await ctx.reply('❌ لم يتم تحديد أي منتج للحذف');
        return;
    }
    
    const orderNumber = session.editingOrder;
    const removedItems = [...session.removeItems];
    
    const success = await removeItemsFromOrder(orderNumber, session.removeItems);
    
    if (success) {
        await ctx.reply(`✅ تم حذف ${session.removeItems.length} منتج من الطلب #${orderNumber} بنجاح`, {
            ...Markup.inlineKeyboard([[Markup.button.callback('🔍 تتبع الطلب', `track_${orderNumber}`)]])
        });
        
        const removedList = removedItems.map(item => `• ${item.quantity} × ${item.name} (${item.price}ج) = ${item.total}ج`).join('\n');
        const adminMsg = `🗑️ <b>تم حذف منتجات من طلب</b>\n━━━━━━━━━━━━━━━\n` +
            `👤 <b>العميل:</b> ${session.name || ctx.from.first_name}\n` +
            `🆔 <b>المعرف:</b> <code>${userId}</code>\n` +
            `🏷️ <b>رقم الطلب:</b> <code>${orderNumber}</code>\n` +
            `━━━━━━━━━━━━━━━\n` +
            `📦 <b>المنتجات المحذوفة:</b>\n${removedList}\n` +
            `━━━━━━━━━━━━━━━\n` +
            `🕐 <b>التاريخ:</b> ${new Date().toLocaleString('ar-EG')}`;
        
        await notifyAdmin(adminMsg);
        
    } else {
        await ctx.reply('❌ حدث خطأ في حذف المنتجات');
    }
    
    session.status = 'main';
    session.editingOrder = null;
    session.removeItems = [];
    session.orderItems = [];
    sessions.set(userId, session);
}

// ============ عرض الإحصائيات ============
async function showStats(ctx) {
    const userId = ctx.from.id;
    const orders = await getUserOrders(userId);
    const totalSpent = await getUserTotalSpent(userId);
    const loyaltyPoints = await getUserLoyaltyPoints(userId);
    const session = sessions.get(userId) || new UserSession(userId);
    session.totalSpent = totalSpent;
    session.loyaltyPoints = loyaltyPoints;
    sessions.set(userId, session);
    
    const pending = orders.filter(o => o.status === 'جاري التنفيذ').length;
    const completed = orders.filter(o => o.status === 'تم التسليم').length;
    const cancelled = orders.filter(o => o.status === 'ملغي').length;
    const totalDiscount = orders.reduce((sum, o) => sum + o.discountAmount, 0);
    
    let tierName = '🟣 جديد', discount = 0, nextAmount = 100;
    if (totalSpent >= 5000) {
        tierName = '💎 بلاتينيوم';
        discount = 10;
        nextAmount = 0;
    } else if (totalSpent >= 2000) {
        tierName = '🌟 ذهبي';
        discount = 7;
        nextAmount = 5000 - totalSpent;
    } else if (totalSpent >= 500) {
        tierName = '⭐ فضي';
        discount = 5;
        nextAmount = 2000 - totalSpent;
    } else if (totalSpent >= 100) {
        tierName = '🟤 برونزي';
        discount = 3;
        nextAmount = 500 - totalSpent;
    } else {
        nextAmount = 100 - totalSpent;
    }
    
    const msg = `📊 <b>إحصائياتي</b>\n━━━━━━━━━━━━━━━\n\n` +
        `👤 <b>الاسم:</b> ${session.name || 'غير مسجل'}\n` +
        `🏆 <b>المستوى:</b> ${tierName} (خصم ${discount}%)\n` +
        `📦 <b>عدد الطلبات:</b> ${orders.length}\n` +
        `✅ <b>تم التسليم:</b> ${completed}\n` +
        `🔄 <b>قيد التنفيذ:</b> ${pending}\n` +
        `❌ <b>ملغي:</b> ${cancelled}\n` +
        `💰 <b>إجمالي المشتريات:</b> ${totalSpent} ج\n` +
        `🎁 <b>إجمالي الخصومات:</b> ${totalDiscount} ج\n` +
        `⭐ <b>نقاط الولاء:</b> ${loyaltyPoints} نقطة\n` +
        (nextAmount > 0 ? `🎯 <b>للمستوى التالي:</b> أنفق ${nextAmount} ج أخرى` : `🏆 <b>أنت في أعلى مستوى!</b>`) +
        `\n\n💡 <b>ملاحظة:</b> كل 10 ج = 1 نقطة ولاء\n` +
        `❌ الطلبات الملغاة لا تحتسب في المشتريات`;
    
    await ctx.reply(msg, {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard([
            [Markup.button.callback('📋 طلباتي', 'my_orders')],
            [Markup.button.callback('🎟️ الكوبونات', 'show_coupons')],
            [Markup.button.callback('⭐ تقييم الخدمة', 'feedback')],
            [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
        ])
    });
}

// ============ التقييم ============
async function showFeedbackButtons(ctx, orderNumber = '') {
    const suffix = orderNumber ? `_${orderNumber}` : '';
    await ctx.reply(
        '⭐ <b>تقييم الخدمة</b>\n\nكيف تقيم تجربتك مع سوبر ماركت الحَواج؟',
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

// ============ تعديل الطلب (إضافة منتجات) ============
async function handleEditOrder(ctx, orderNumber) {
    const order = await getOrderByNumber(orderNumber);
    
    if (!order || !order.canEdit) {
        await ctx.reply('❌ لا يمكن تعديل هذا الطلب حالياً');
        return;
    }
    
    const session = sessions.get(ctx.from.id) || new UserSession(ctx.from.id);
    session.status = 'editing';
    session.editingOrder = orderNumber;
    session.editCart = [];
    sessions.set(ctx.from.id, session);
    
    await ctx.reply(`✏️ <b>إضافة منتجات إلى الطلب #${orderNumber}</b>\n\n` +
        `📦 المنتجات الحالية:\n${order.itemsText}\n\n` +
        `➕ أضف منتجات جديدة بالضغط على "🛒 المنتجات" أو "📂 الأقسام"\n` +
        `المنتجات الجديدة ستضاف إلى الطلب الحالي\n\n` +
        `بعد الانتهاء، اضغط "إنهاء التعديل"`,
        {
            parse_mode: 'HTML',
            ...Markup.inlineKeyboard([
                [Markup.button.callback('🛒 المنتجات', 'browse_products')],
                [Markup.button.callback('📂 الأقسام', 'show_categories')],
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
        await ctx.reply('⚠️ لم تقم بإضافة أي منتجات جديدة\n\nلإضافة منتجات، اضغط على "🛒 المنتجات" أو "📂 الأقسام" ثم اختر المنتجات');
        return;
    }
    
    const orderNumber = session.editingOrder;
    const addedItems = [...session.editCart];
    
    const success = await addItemsToOrder(orderNumber, session.editCart);
    
    if (success) {
        const addedItemsText = addedItems.map(item => `${item.quantity} × ${item.name}`).join(', ');
        
        session.editCart = [];
        session.status = 'main';
        session.editingOrder = null;
        sessions.set(userId, session);
        
        await ctx.reply(`✅ تم إضافة منتجات إلى الطلب #${orderNumber} بنجاح!\n\n📦 تم إضافة: ${addedItemsText}`, {
            ...Markup.inlineKeyboard([[Markup.button.callback('🔍 تتبع الطلب', `track_${orderNumber}`)]])
        });
        
        const addedList = addedItems.map(item => `• ${item.quantity} × ${item.name} (${item.price}ج) = ${item.price * item.quantity}ج`).join('\n');
        const adminMsg = `➕ <b>تم إضافة منتجات إلى طلب</b>\n━━━━━━━━━━━━━━━\n` +
            `👤 <b>العميل:</b> ${session.name || 'عميل'}\n` +
            `🆔 <b>المعرف:</b> <code>${userId}</code>\n` +
            `🏷️ <b>رقم الطلب:</b> <code>${orderNumber}</code>\n` +
            `━━━━━━━━━━━━━━━\n` +
            `📦 <b>المنتجات المضافة:</b>\n${addedList}\n` +
            `━━━━━━━━━━━━━━━\n` +
            `🕐 <b>التاريخ:</b> ${new Date().toLocaleString('ar-EG')}`;
        
        await notifyAdmin(adminMsg);
        
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
        session.removeItems = [];
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
        await ctx.reply('❌ لا يمكن إلغاء هذا الطلب حالياً');
        return;
    }
    
    const session = sessions.get(ctx.from.id) || new UserSession(ctx.from.id);
    session.tempOrderNumber = orderNumber;
    sessions.set(ctx.from.id, session);
    
    await ctx.reply(`⚠️ <b>تأكيد إلغاء الطلب</b>\n\n` +
        `🏷️ ${order.orderNumber}\n` +
        `💰 ${order.finalTotal} ج\n` +
        `⭐ نقاط الولاء: ${order.loyaltyPoints} نقطة\n\n` +
        `⚠️ <b>ملاحظة مهمة:</b>\n` +
        `• سيتم استرجاع نقاط الولاء\n` +
        `• سيتم خصم قيمة الطلب من إجمالي مشترياتك\n` +
        `• قد ينخفض مستواك\n\n` +
        `هل أنت متأكد من إلغاء هذا الطلب؟`,
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
    
    const order = await getOrderByNumber(session.tempOrderNumber);
    
    if (order) {
        const success = await updateOrderStatus(session.tempOrderNumber, 'ملغي');
        
        if (success) {
            const newTotalSpent = await getUserTotalSpent(userId);
            session.totalSpent = newTotalSpent;
            session.subtractLoyaltyPoints(order.finalTotal);
            sessions.set(userId, session);
            
            const newTier = session.getTier();
            
            await ctx.reply(
                `✅ <b>تم إلغاء الطلب #${session.tempOrderNumber} بنجاح</b>\n\n` +
                `📊 <b>تحديث بياناتك:</b>\n` +
                `💰 إجمالي المشتريات: ${newTotalSpent} ج\n` +
                `⭐ نقاط الولاء: ${session.loyaltyPoints}\n` +
                `🏆 مستواك الجديد: ${newTier.name} (خصم ${newTier.discount}%)\n\n` +
                `🛒 يمكنك تقديم طلب جديد الآن`,
                {
                    parse_mode: 'HTML',
                    ...Markup.inlineKeyboard([
                        [Markup.button.callback('🛒 طلب جديد', 'browse_products')],
                        [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
                    ])
                }
            );
            
            const adminMsg = `❌ <b>تم إلغاء طلب</b>\n━━━━━━━━━━━━━━━\n` +
                `👤 <b>العميل:</b> ${session.name || 'عميل'}\n` +
                `🆔 <b>المعرف:</b> <code>${userId}</code>\n` +
                `🏷️ <b>رقم الطلب:</b> <code>${session.tempOrderNumber}</code>\n` +
                `💰 <b>قيمة الطلب:</b> ${order.finalTotal} ج\n` +
                `⭐ <b>نقاط الولاء:</b> ${order.loyaltyPoints} نقطة\n` +
                `📅 <b>تاريخ الطلب:</b> ${order.date}\n` +
                `━━━━━━━━━━━━━━━\n` +
                `🕐 <b>تاريخ الإلغاء:</b> ${new Date().toLocaleString('ar-EG')}`;
            
            await notifyAdmin(adminMsg);
            
        } else {
            await ctx.reply('❌ لا يمكن إلغاء هذا الطلب الآن');
        }
    } else {
        await ctx.reply('❌ لم يتم العثور على الطلب');
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
    
    const subtotal = cart.reduce((sum, item) => sum + item.total, 0);
    const session = sessions.get(userId) || new UserSession(userId);
    const discounts = await calculateDiscounts(userId, subtotal, session.appliedCoupon);
    
    const lastData = await getUserLastAddress(userId);
    
    let cartSummary = '';
    for (let i = 0; i < cart.length; i++) {
        cartSummary += `${i+1}. ${cart[i].quantity} × ${cart[i].name} = ${cart[i].total} ج\n`;
    }
    
    let discountSummary = '';
    if (discounts.tierDiscount.amount > 0 && discounts.tierDiscount.applied) {
        discountSummary += `🏆 خصم المستوى (${discounts.tierDiscount.percent}%): -${discounts.tierDiscount.amount} ج\n`;
    }
    if (discounts.couponDiscount.valid && discounts.couponDiscount.applied && discounts.couponDiscount.amount > 0) {
        discountSummary += `🎟️ خصم الكوبون (${discounts.couponDiscount.percent}%): -${discounts.couponDiscount.amount} ج\n`;
    }
    
    const expectedPoints = Math.floor(discounts.finalTotal / 10);
    
    if (lastData && session.useSavedAddress === null) {
        await ctx.reply(
            `🛍️ <b>تأكيد الطلب</b>\n━━━━━━━━━━━━━━━\n\n` +
            cartSummary +
            `\n💰 <b>المجموع:</b> ${discounts.subtotal} ج\n` +
            discountSummary +
            `━━━━━━━━━━━━━━━\n` +
            `💎 <b>الإجمالي بعد الخصم:</b> ${discounts.finalTotal} ج\n` +
            `⭐ <b>نقاط الولاء المتوقعة:</b> ${expectedPoints} نقطة\n━━━━━━━━━━━━━━━\n\n` +
            `📦 <b>هل تريد استخدام بياناتك السابقة؟</b>\n\n` +
            `📍 العنوان: ${lastData.address}\n` +
            `📞 الهاتف: ${lastData.phone}`,
            {
                parse_mode: 'HTML',
                ...Markup.inlineKeyboard([
                    [Markup.button.callback('✅ نعم، استخدم البيانات', 'use_saved_data')],
                    [Markup.button.callback('📝 لا، أدخل بيانات جديدة', 'enter_new_data')],
                    [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
                ])
            }
        );
        return;
    }
    
    if (lastData && session.useSavedAddress === true) {
        session.address = lastData.address;
        session.phone = lastData.phone;
        session.useSavedAddress = null;
        sessions.set(userId, session);
        await saveOrderWithData(ctx, userId, cart, discounts);
        return;
    }
    
    session.useSavedAddress = null;
    session.status = 'ordering_name';
    sessions.set(userId, session);
    
    await ctx.reply(
        `🛍️ <b>تأكيد الطلب</b>\n━━━━━━━━━━━━━━━\n\n${cartSummary}\n` +
        `💰 <b>المجموع:</b> ${discounts.subtotal} ج\n` +
        discountSummary +
        `━━━━━━━━━━━━━━━\n` +
        `💎 <b>الإجمالي بعد الخصم:</b> ${discounts.finalTotal} ج\n` +
        `⭐ <b>نقاط الولاء المتوقعة:</b> ${expectedPoints} نقطة\n━━━━━━━━━━━━━━━\n\n` +
        `📝 أرسل اسمك الكامل:`,
        { parse_mode: 'HTML' }
    );
}

async function saveOrderWithData(ctx, userId, cart, discounts) {
    const session = sessions.get(userId);
    const orderCart = [...cart];
    const loyaltyPoints = Math.floor(discounts.finalTotal / 10);
    
    const saved = await saveOrder(
        userId, session.name || 'عميل', session.address, session.phone, cart,
        discounts.subtotal, discounts.finalTotal, loyaltyPoints
    );
    
    if (saved.success) {
        const pointsEarned = session.addLoyaltyPoints(discounts.finalTotal);
        session.totalSpent += discounts.finalTotal;
        
        if (discounts.usedCoupon) {
            await markCouponAsUsed(userId, discounts.usedCoupon);
            session.removeCoupon();
        }
        
        const tier = session.getTier();
        
        carts.set(userId, []);
        session.status = 'main';
        sessions.set(userId, session);
        
        const productList = orderCart.map((item, i) => `${i+1}. ${item.quantity} × ${item.name} = ${item.total} ج`).join('\n');
        
        let discountText = '';
        if (discounts.tierDiscount.amount > 0 && discounts.tierDiscount.applied) {
            discountText += `🏆 <b>خصم المستوى (${discounts.tierDiscount.percent}%):</b> -${discounts.tierDiscount.amount} ج\n`;
        }
        if (discounts.couponDiscount.valid && discounts.couponDiscount.applied && discounts.couponDiscount.amount > 0) {
            discountText += `🎟️ <b>خصم الكوبون (${discounts.couponDiscount.percent}%):</b> -${discounts.couponDiscount.amount} ج\n`;
        }
        
        const successMessage = 
`✅ <b>تم تسجيل طلبك بنجاح!</b>
━━━━━━━━━━━━━━━

🏷️ <b>رقم الطلب:</b> <code>${saved.orderNumber}</code>

👤 <b>الاسم:</b> ${session.name}
━━━━━━━━━━━━━━━
📦 <b>المنتجات:</b>
${productList}
━━━━━━━━━━━━━━━
💰 <b>المجموع:</b> ${discounts.subtotal} ج
${discountText}
━━━━━━━━━━━━━━━
💎 <b>الإجمالي المدفوع:</b> ${discounts.finalTotal} ج
━━━━━━━━━━━━━━━
⭐ <b>نقاط الولاء المكتسبة:</b> ${pointsEarned} نقطة
━━━━━━━━━━━━━━━

📍 <b>العنوان:</b> ${session.address}
📞 <b>الهاتف:</b> ${session.phone}

🏆 <b>مستواك الحالي:</b> ${tier.name} (خصم ${tier.discount}%)

📌 <b>لتتبع طلبك:</b>
اضغط على 🔍 تتبع طلب في القائمة وأرسل: <code>${saved.orderNumber}</code>

🙏 شكراً لتسوقك مع سوبر ماركت الحَواج!`;
        
        await ctx.reply(successMessage, {
            parse_mode: 'HTML',
            ...Markup.inlineKeyboard([
                [Markup.button.callback('🔍 تتبع هذا الطلب', `track_${saved.orderNumber}`)],
                [Markup.button.callback('⭐ تقييم الخدمة', 'feedback')],
                [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
            ])
        });
        
        const adminMsg = `🛍️ <b>طلب جديد!</b>\n━━━━━━━━━━━━━━━\n` +
            `👤 <b>العميل:</b> ${session.name}\n` +
            `🆔 <b>المعرف:</b> <code>${userId}</code>\n` +
            `📱 <b>اسم المستخدم:</b> @${ctx.from.username || 'لا يوجد'}\n` +
            `🏷️ <b>رقم الطلب:</b> <code>${saved.orderNumber}</code>\n` +
            `━━━━━━━━━━━━━━━\n` +
            `📦 <b>المنتجات:</b>\n${productList}\n` +
            `━━━━━━━━━━━━━━━\n` +
            `💰 <b>المجموع:</b> ${discounts.subtotal} ج\n` +
            `${discountText}` +
            `💎 <b>الإجمالي المدفوع:</b> ${discounts.finalTotal} ج\n` +
            `⭐ <b>نقاط الولاء:</b> ${pointsEarned} نقطة\n` +
            `━━━━━━━━━━━━━━━\n` +
            `📍 <b>العنوان:</b> ${session.address}\n` +
            `📞 <b>الهاتف:</b> ${session.phone}\n` +
            `━━━━━━━━━━━━━━━\n` +
            `🏆 <b>مستوى العميل:</b> ${tier.name} (خصم ${tier.discount}%)\n` +
            `━━━━━━━━━━━━━━━\n` +
            `🕐 <b>تاريخ الطلب:</b> ${new Date().toLocaleString('ar-EG')}`;
        
        await notifyAdmin(adminMsg);
        
    } else {
        await ctx.reply('❌ حدث خطأ في حفظ الطلب، حاول مرة أخرى');
    }
}

// ============ البحث والتتبع ============
async function handleSearch(ctx) {
    await ctx.reply('🔍 اكتب اسم المنتج:', {
        ...Markup.inlineKeyboard([
            [Markup.button.callback('📂 الأقسام', 'show_categories')],
            [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
        ])
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
    [Markup.button.callback('📂 الأقسام', 'show_categories')],
    [Markup.button.callback('🛍️ السلة', 'view_cart')],
    [Markup.button.callback('🎁 العروض', 'view_offers')],
    [Markup.button.callback('🎟️ الكوبونات', 'show_coupons')],
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
    
    const totalSpent = await getUserTotalSpent(userId);
    const loyaltyPoints = await getUserLoyaltyPoints(userId);
    session.totalSpent = totalSpent;
    session.loyaltyPoints = loyaltyPoints;
    const tier = session.getTier();
    
    const welcomeMsg = `🌟 <b>مرحباً بك في سوبر ماركت الحَواج!</b> 🌟\n━━━━━━━━━━━━━━━\n\n` +
        `👤 <b>الاسم:</b> ${ctx.from.first_name}\n` +
        `🏆 <b>مستواك:</b> ${tier.name} (خصم ${tier.discount}%)\n` +
        `⭐ <b>نقاط الولاء:</b> ${loyaltyPoints} نقطة\n\n` +
        `🛒 أكبر سوبر ماركت في العالم!\n` +
        `✅ توصيل سريع - أسعار تنافسية - جودة عالية\n` +
        `📂 تصفح المنتجات حسب الأقسام!\n` +
        `💡 <b>ملاحظة:</b> كل 10 ج = 1 نقطة ولاء\n━━━━━━━━━━━━━━━\n\n` +
        `📌 اختر من القائمة:`;
    
    await ctx.reply(welcomeMsg, { parse_mode: 'HTML', ...mainKeyboard });
});

bot.action('main_menu', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await ctx.reply('🏠 القائمة الرئيسية', { ...mainKeyboard });
});

bot.action('browse_products', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    currentCategory = '';
    await showProducts(ctx);
});

bot.action('show_categories', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await showCategories(ctx);
});

bot.action(/category_(.+)/, async (ctx) => {
    const category = ctx.match[1];
    await ctx.answerCbQuery(`📂 عرض منتجات قسم ${category}`).catch(() => {});
    currentCategory = category;
    await showProducts(ctx, 0, '', category);
});

bot.action('view_cart', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await showCart(ctx);
});

bot.action('view_offers', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await showOffers(ctx);
});

bot.action('show_coupons', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await showAvailableCoupons(ctx);
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
    await ctx.reply(
        '📞 <b>الدعم الفني</b>\n━━━━━━━━━━━━━━━\n\n' +
        '📧 <b>الإيميل:</b> <code>muhammedhosni70@gmail.com</code>\n' +
        '📱 <b>واتساب:</b> <code>01020063819</code>\n' +
        '✈️ <b>تليجرام:</b> @Muhamedhosny\n\n' +
        '⏰ <b>خدمة عملاء 24 ساعة</b>',
        { parse_mode: 'HTML' }
    );
});

bot.action('search_products', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await handleSearch(ctx);
});

bot.action('enter_coupon', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await enterCoupon(ctx);
});

bot.action('next_page', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await showProducts(ctx, currentPage + 1, currentSearch, currentCategory);
});

bot.action('prev_page', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await showProducts(ctx, currentPage - 1, currentSearch, currentCategory);
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
    
    // البحث عن القسم للمنتج
    const products = await getProducts();
    const product = products.find(p => p.name === productName);
    const category = product ? product.category : '';
    
    if (session && session.status === 'editing') {
        const existing = session.editCart.find(item => item.name === productName);
        if (existing) {
            existing.quantity++;
        } else {
            session.editCart.push({ name: productName, price: productPrice, quantity: 1 });
        }
        sessions.set(userId, session);
        await ctx.reply(`✅ تم إضافة ${productName} إلى تعديل الطلب`);
    } else {
        let cart = carts.get(userId) || [];
        const existing = cart.find(item => item.name === productName);
        if (existing) {
            existing.quantity++;
        } else {
            cart.push(new CartItem(productName, productPrice, 1, category));
        }
        carts.set(userId, cart);
        await ctx.reply(`✅ تم إضافة ${productName} إلى السلة`);
    }
});

bot.action(/order_offer_(\d+)/, async (ctx) => {
    const offerIndex = ctx.match[1];
    await ctx.answerCbQuery().catch(() => {});
    await handleOrderOffer(ctx, offerIndex);
});

bot.action(/edit_order_(.+)/, async (ctx) => {
    const orderNumber = ctx.match[1];
    await ctx.answerCbQuery().catch(() => {});
    await handleEditOrder(ctx, orderNumber);
});

bot.action(/delete_items_(.+)/, async (ctx) => {
    const orderNumber = ctx.match[1];
    await ctx.answerCbQuery().catch(() => {});
    await showOrderItemsForDeletion(ctx, orderNumber);
});

bot.action(/remove_item_(\d+)/, async (ctx) => {
    const itemIndex = ctx.match[1];
    await ctx.answerCbQuery().catch(() => {});
    await markItemForRemoval(ctx, itemIndex);
});

bot.action('confirm_remove_items', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
    await confirmRemoveItems(ctx);
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

bot.action('noop', async (ctx) => {
    await ctx.answerCbQuery().catch(() => {});
});

bot.action('use_saved_data', async (ctx) => {
    const userId = ctx.from.id;
    const session = sessions.get(userId) || new UserSession(userId);
    session.useSavedAddress = true;
    sessions.set(userId, session);
    await ctx.answerCbQuery().catch(() => {});
    await startCheckout(ctx);
});

bot.action('enter_new_data', async (ctx) => {
    const userId = ctx.from.id;
    const session = sessions.get(userId) || new UserSession(userId);
    session.useSavedAddress = false;
    sessions.set(userId, session);
    await ctx.answerCbQuery().catch(() => {});
    await startCheckout(ctx);
});

// ============ أزرار التقييم ============
for (let i = 1; i <= 5; i++) {
    bot.action(new RegExp(`rate_${i}(?:_(.*))?`), async (ctx) => {
        const rating = i;
        const userId = ctx.from.id;
        const session = sessions.get(userId) || new UserSession(userId);
        
        await ctx.answerCbQuery(`✅ شكراً لتقييمك ${rating} نجوم!`).catch(() => {});
        
        await saveFeedback(userId, session.name || ctx.from.first_name, rating);
        
        await ctx.reply(
            `🙏 <b>شكراً لتقييمك ${'⭐'.repeat(rating)}</b>\n\n` +
            `نقدر رأيك ونسعى دائماً لتحسين خدماتنا.`,
            {
                parse_mode: 'HTML',
                ...Markup.inlineKeyboard([
                    [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
                ])
            }
        );
        
        const adminMsg = `⭐ <b>تقييم جديد</b>\n━━━━━━━━━━━━━━━\n` +
            `👤 <b>العميل:</b> ${session.name || ctx.from.first_name}\n` +
            `🆔 <b>المعرف:</b> <code>${userId}</code>\n` +
            `📱 <b>اسم المستخدم:</b> @${ctx.from.username || 'لا يوجد'}\n` +
            `━━━━━━━━━━━━━━━\n` +
            `🎖️ <b>التقييم:</b> ${rating} / 5 ${'⭐'.repeat(rating)}\n` +
            `━━━━━━━━━━━━━━━\n` +
            `🕐 <b>التاريخ:</b> ${new Date().toLocaleString('ar-EG')}`;
        
        await notifyAdmin(adminMsg);
    });
}

bot.action(/rate_order_(.+)/, async (ctx) => {
    const orderNumber = ctx.match[1];
    await ctx.answerCbQuery().catch(() => {});
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
    
    const lastMessage = userCooldowns.get(userId);
    if (lastMessage && Date.now() - lastMessage < 1000) return;
    userCooldowns.set(userId, Date.now());
    
    if (session.status === 'entering_coupon') {
        const couponCode = text.toUpperCase().trim();
        const couponCheck = await validateCoupon(couponCode, userId);
        
        if (couponCheck.valid) {
            session.applyCoupon(couponCode, couponCheck.discount);
            session.status = 'main';
            sessions.set(userId, session);
            
            await ctx.reply(`✅ <b>تم تطبيق الكوبون بنجاح!</b>\n\n🎟️ ${couponCheck.message}`, {
                parse_mode: 'HTML',
                ...Markup.inlineKeyboard([
                    [Markup.button.callback('🛍️ عرض السلة', 'view_cart')],
                    [Markup.button.callback('🛒 متابعة التسوق', 'browse_products')]
                ])
            });
        } else {
            await ctx.reply(`${couponCheck.message}`, {
                ...Markup.inlineKeyboard([
                    [Markup.button.callback('🎟️ إدخال كوبون آخر', 'enter_coupon')],
                    [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
                ])
            });
        }
        return;
    }
    
    if (session.status === 'tracking') {
        session.status = 'main';
        sessions.set(userId, session);
        await trackOrder(ctx, text);
        return;
    }
    
    if (session.status === 'searching') {
        session.status = 'main';
        sessions.set(userId, session);
        await showProducts(ctx, 0, text, '');
        return;
    }
    
    if (session.status === 'ordering_name') {
        if (text.length < 3) {
            await ctx.reply('❌ الاسم قصير جداً، أرسل اسمك الكامل:');
            return;
        }
        session.name = text;
        session.status = 'ordering_address';
        sessions.set(userId, session);
        await ctx.reply(`✅ تم حفظ: ${text}\n\n📍 أرسل عنوانك:`);
        return;
    }
    
    if (session.status === 'ordering_address') {
        if (text.length < 10) {
            await ctx.reply('❌ العنوان غير مفصل، أرسل عنواناً مفصلاً:');
            return;
        }
        session.address = text;
        session.status = 'ordering_phone';
        sessions.set(userId, session);
        await ctx.reply(`✅ تم حفظ العنوان\n\n📞 أرسل رقم هاتفك:`);
        return;
    }
    
    if (session.status === 'ordering_phone') {
        const phoneMatch = text.match(/01[0125][0-9]{8}/);
        if (!phoneMatch) {
            await ctx.reply('❌ رقم غير صحيح، أرسل رقم مصري 11 رقم يبدأ بـ 01:');
            return;
        }
        
        session.phone = phoneMatch[0];
        const cart = carts.get(userId) || [];
        const subtotal = cart.reduce((sum, item) => sum + item.total, 0);
        const discounts = await calculateDiscounts(userId, subtotal, session.appliedCoupon);
        
        const saved = await saveOrder(
            userId, session.name, session.address, session.phone, cart,
            subtotal, discounts.finalTotal, Math.floor(discounts.finalTotal / 10)
        );
        
        if (saved.success) {
            const pointsEarned = session.addLoyaltyPoints(discounts.finalTotal);
            session.totalSpent += discounts.finalTotal;
            
            if (discounts.usedCoupon) {
                await markCouponAsUsed(userId, discounts.usedCoupon);
                session.removeCoupon();
            }
            
            const tier = session.getTier();
            
            carts.set(userId, []);
            session.status = 'main';
            sessions.set(userId, session);
            
            const productList = cart.map((item, i) => `${i+1}. ${item.quantity} × ${item.name} = ${item.total} ج`).join('\n');
            
            let discountText = '';
            if (discounts.tierDiscount.amount > 0 && discounts.tierDiscount.applied) {
                discountText += `🏆 <b>خصم المستوى (${discounts.tierDiscount.percent}%):</b> -${discounts.tierDiscount.amount} ج\n`;
            }
            if (discounts.couponDiscount.valid && discounts.couponDiscount.applied && discounts.couponDiscount.amount > 0) {
                discountText += `🎟️ <b>خصم الكوبون (${discounts.couponDiscount.percent}%):</b> -${discounts.couponDiscount.amount} ج\n`;
            }
            
            const successMessage = 
`✅ <b>تم تسجيل طلبك بنجاح!</b>
━━━━━━━━━━━━━━━

🏷️ <b>رقم الطلب:</b> <code>${saved.orderNumber}</code>

👤 <b>الاسم:</b> ${session.name}
━━━━━━━━━━━━━━━
📦 <b>المنتجات:</b>
${productList}
━━━━━━━━━━━━━━━
💰 <b>المجموع:</b> ${subtotal} ج
${discountText}
━━━━━━━━━━━━━━━
💎 <b>الإجمالي المدفوع:</b> ${discounts.finalTotal} ج
━━━━━━━━━━━━━━━
⭐ <b>نقاط الولاء المكتسبة:</b> ${pointsEarned} نقطة
━━━━━━━━━━━━━━━

📍 <b>العنوان:</b> ${session.address}
📞 <b>الهاتف:</b> ${session.phone}

🏆 <b>مستواك الحالي:</b> ${tier.name} (خصم ${tier.discount}%)

📌 <b>لتتبع طلبك:</b>
اضغط على 🔍 تتبع طلب في القائمة وأرسل: <code>${saved.orderNumber}</code>

🙏 شكراً لتسوقك مع سوبر ماركت الحَواج!`;
            
            await ctx.reply(successMessage, {
                parse_mode: 'HTML',
                ...Markup.inlineKeyboard([
                    [Markup.button.callback('🔍 تتبع هذا الطلب', `track_${saved.orderNumber}`)],
                    [Markup.button.callback('⭐ تقييم الخدمة', 'feedback')],
                    [Markup.button.callback('🏠 الرئيسية', 'main_menu')]
                ])
            });
            
            const adminMsg = `🛍️ <b>طلب جديد!</b>\n━━━━━━━━━━━━━━━\n` +
                `👤 <b>العميل:</b> ${session.name}\n` +
                `🆔 <b>المعرف:</b> <code>${userId}</code>\n` +
                `📱 <b>اسم المستخدم:</b> @${ctx.from.username || 'لا يوجد'}\n` +
                `🏷️ <b>رقم الطلب:</b> <code>${saved.orderNumber}</code>\n` +
                `━━━━━━━━━━━━━━━\n` +
                `📦 <b>المنتجات:</b>\n${productList}\n` +
                `━━━━━━━━━━━━━━━\n` +
                `💰 <b>المجموع:</b> ${subtotal} ج\n` +
                `${discountText}` +
                `💎 <b>الإجمالي المدفوع:</b> ${discounts.finalTotal} ج\n` +
                `⭐ <b>نقاط الولاء:</b> ${pointsEarned} نقطة\n` +
                `━━━━━━━━━━━━━━━\n` +
                `📍 <b>العنوان:</b> ${session.address}\n` +
                `📞 <b>الهاتف:</b> ${session.phone}\n` +
                `━━━━━━━━━━━━━━━\n` +
                `🏆 <b>مستوى العميل:</b> ${tier.name} (خصم ${tier.discount}%)\n` +
                `━━━━━━━━━━━━━━━\n` +
                `🕐 <b>تاريخ الطلب:</b> ${new Date().toLocaleString('ar-EG')}`;
            
            await notifyAdmin(adminMsg);
            
        } else {
            await ctx.reply('❌ حدث خطأ في حفظ الطلب، حاول مرة أخرى');
        }
        return;
    }
    
    if (text.length > 2 && !text.startsWith('/')) {
        await showProducts(ctx, 0, text, '');
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
    await getCoupons(true);
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

// ============ Webhook لـ Vercel ============
export default async function handler(req, res) {
    try {
        if (req.method === 'POST') {
            await bot.handleUpdate(req.body);
            res.status(200).json({ status: 'ok' });
        } else {
            res.status(200).send('✅ سوبر ماركت الحَواج يعمل بنجاح!');
        }
    } catch (error) {
        console.error('❌ Webhook error:', error);
        res.status(500).json({ error: error.message });
    }
}

if (process.env.NODE_ENV !== 'production') {
    bot.launch();
    console.log('🤖 سوبر ماركت الحَواج يعمل محلياً...');
}

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
