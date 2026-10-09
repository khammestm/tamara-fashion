/**
 * =========================================================================
 * TAMARA FASHION - Google Apps Script Backend (Code.gs)
 * =========================================================================
 * هذا السكربت يقوم بحفظ طلبات متجر TAMARA FASHION تلقائيًا في Google Sheet
 * مع الحماية الكاملة من السبام وتكرار الطلبات وتأمين الأسعار وحقن المعادلات.
 * 
 * المطور لـ: TAMARA FASHION
 * المنتوج: فستان ميرال قطيفة طقم كامل 3 قطع (بيج، عنابي، بلون كاسي)
 * تفصيل ورشة تمارا في بوسعادة منذ 2020
 * =========================================================================
 */

// إعدادات المتجر على السيرفر (لا يمكن للعميل أو المتصفح التلاعب بها)
const SERVER_CONFIG = {
  SHEET_NAME: "Orders",
  // معرف جدول Google Sheet (اختياري):
  // - اتركه فارغًا "" إذا فتحت السكربت من داخل الشيت عبر: Extensions > Apps Script
  // - أو ضع معرف الجدول (Spreadsheet ID) هنا إذا كان السكربت مستقلاً
  SPREADSHEET_ID: "",
  
  PRODUCT_NAME: "فستان ميرال قطيفة طقم كامل 3 قطع",
  UNIT_PRICE: 4900, // السعر الثابت 4900 دج
  DELIVERY_FEES: {
    domicile: 700, // توصيل للمنزل: 700 دج
    yalidine: 550  // استلام من مكتب ياليدين: 550 دج
  },
  ALLOWED_SIZES: ["36", "38", "40", "42", "44", "46", "48", "50"],
  ALLOWED_COLORS: ["بيج (Beige)", "عنابي (Bordeaux)", "بلون كاسي (Blanc Cassé)"],
  STATUS_OPTIONS: ["جديد", "مؤكد", "قيد التحضير", "تم الشحن", "تم التسليم", "ملغي"],
  DEFAULT_STATUS: "جديد",
  TIMEZONE: "Africa/Algiers"
};

/**
 * دالة الحصول على ملف جدول Google Sheet المستهدف
 */
function getTargetSpreadsheet() {
  if (SERVER_CONFIG.SPREADSHEET_ID && SERVER_CONFIG.SPREADSHEET_ID.trim() !== "") {
    try {
      return SpreadsheetApp.openById(SERVER_CONFIG.SPREADSHEET_ID.trim());
    } catch (e) {
      Logger.log("تعذر فتح الجدول بالمعرف المحدد: " + e.message);
    }
  }
  
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  
  throw new Error("لم يتم العثور على جدول Google Sheet. إذا كان السكربت مستقلاً، يرجى كتابة SPREADSHEET_ID في كائن SERVER_CONFIG أو فتح الشيت والضغط على Extensions > Apps Script.");
}

/**
 * دالة الإعداد التلقائي بضغطة زر واحدة (Run setupSheet)
 * تقوم بتهيئة الورقة "Orders"، إضافة العناوين، تجميد السطر الأول، 
 * وتعيين القوائم المنسدلة لحالة الطلب مع ألوان التمييز التلقائية.
 */
function setupSheet() {
  const ss = getTargetSpreadsheet();
  let sheet = ss.getSheetByName(SERVER_CONFIG.SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SERVER_CONFIG.SHEET_NAME);
  }
  
  // عناوين الأعمدة المطلوبة بالترتيب الدقيق
  const headers = [
    "Order ID",
    "Date/Time",
    "Status",
    "Full Name",
    "Phone",
    "City",
    "Address",
    "Product",
    "Size",
    "Quantity",
    "Unit Price",
    "Delivery Fee",
    "Total",
    "Notes",
    "My Notes"
  ];
  
  // ضبط العناوين
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  
  // تنسيق سطر العناوين (لون عنابي فاخر يليق بماركة TAMARA FASHION)
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#6B1426");
  headerRange.setFontColor("#FFFFFF");
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");
  sheet.setRowHeight(1, 40);
  
  // إنشاء القائمة المنسدلة لعمود الحالة (Status - العمود C)
  const statusColumnRange = sheet.getRange("C2:C5000");
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(SERVER_CONFIG.STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  statusColumnRange.setDataValidation(rule);
  
  // إعداد التنسيق الشرطي التلقائي لألوان الحالات
  const rules = [];
  const colorMap = [
    { value: "جديد", bg: "#FFF3CD", text: "#856404" },         // أصفر دافئ
    { value: "مؤكد", bg: "#CCE5FF", text: "#004085" },         // أزرق هادئ
    { value: "قيد التحضير", bg: "#E2D9F3", text: "#49128A" },  // بنفسجي
    { value: "تم الشحن", bg: "#D1ECF1", text: "#0C5460" },      // تركواز
    { value: "تم التسليم", bg: "#D4EDDA", text: "#155724" },    // أخضر نجاح
    { value: "ملغي", bg: "#F8D7DA", text: "#721C24" }          // أحمر إلغاء
  ];
  
  colorMap.forEach(item => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(item.value)
        .setBackground(item.bg)
        .setFontColor(item.text)
        .setBold(true)
        .setRanges([statusColumnRange])
        .build()
    );
  });
  
  sheet.setConditionalFormatRules(rules);
  
  // محاذاة وتنسيق افتراضي للأعمدة
  sheet.getRange("A2:A5000").setHorizontalAlignment("center"); // Order ID
  sheet.getRange("B2:B5000").setHorizontalAlignment("center"); // Date/Time
  sheet.getRange("C2:C5000").setHorizontalAlignment("center"); // Status
  sheet.getRange("E2:E5000").setHorizontalAlignment("center"); // Phone
  sheet.getRange("I2:I5000").setHorizontalAlignment("center"); // Size
  sheet.getRange("J2:J5000").setHorizontalAlignment("center"); // Quantity
  sheet.getRange("K2:M5000").setHorizontalAlignment("right");  // Prices (DA)
  
  // تنسيق العملة
  sheet.getRange("K2:M5000").setNumberFormat('#,##0" دج"');
  
  SpreadsheetApp.flush();
  Logger.log("تمت تهيئة ورقة Orders بنجاح وتعيين القوائم والتنسيقات!");
}

/**
 * فحص الاتصال البسيط (GET Request) أو استقبال الطلبات عبر GET
 */
function doGet(e) {
  // فحص ما إذا كان هناك طلب شراء مرسل عبر رابط GET (حل احتياطي قوي ومضمون)
  if (e && e.parameter && (e.parameter.phone || e.parameter.fullName || e.parameter.city)) {
    return handleOrderProcessing(e);
  }
  
  // فحص الاتصال الصحي (Health Check)
  return createJsonResponse({
    status: "ok",
    success: true,
    message: "TAMARA FASHION Order API is running perfectly.",
    access: "Anyone (Public)",
    timestamp: new Date().toISOString()
  });
}

/**
 * استقبال ومعالجة طلبات الشراء الجديدة (POST Request)
 */
function doPost(e) {
  return handleOrderProcessing(e);
}

/**
 * المعالج الرئيسي المشترك لحفظ الطلبات
 */
function handleOrderProcessing(e) {
  // استخدام LockService لضمان عدم تضارب المعرفات عند قدوم طلبين في نفس اللحظة
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(15000); // الانتظار حتى 15 ثانية
  
  if (!hasLock) {
    return createJsonResponse({
      success: false,
      error: "الخادم مشغول حاليًا، يرجى إعادة المحاولة بعد ثوانٍ قليلـة."
    });
  }
  
  try {
    let data = {};
    
    // استخراج البيانات سواء تم إرسالها بتنسيق JSON أو Form-urlencoded
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        if (e.parameter) {
          data = Object.assign({}, e.parameter);
        }
      }
    }
    
    if (e && e.parameter) {
      data = Object.assign({}, e.parameter, data);
    }

    // 1. فحص مصيدة الروبوتات (Honeypot): إذا امتلأ الحقل المخفي يتم تجاهل الطلب
    if (data.website || data.honeypot || data.middle_name) {
      return createJsonResponse({
        success: true,
        orderId: "TF-BOT",
        message: "تم استلام طلبك، سنتصل بك قريبًا لتأكيده"
      });
    }

    // 2. تنظيف والتحقق من الحقول الإلزامية
    const fullName = sanitizeInput(data.fullName || data.name || "");
    const phone = sanitizeInput(data.phone || "");
    const city = sanitizeInput(data.city || data.wilaya || "");
    const address = sanitizeInput(data.address || "");
    let size = String(data.size || "40").trim();
    const color = sanitizeInput(data.color || "عنابي (Bordeaux)");
    let quantity = parseInt(data.quantity, 10);
    const deliveryType = String(data.deliveryType || "domicile").toLowerCase().trim();
    const notes = sanitizeInput(data.notes || "");

    if (!fullName || fullName.length < 2 || fullName.length > 100) {
      return createJsonResponse({ success: false, error: "يرجى كتابة الاسم الكامل بشكل صحيح." });
    }

    // التحقق من صحة رقم الهاتف الجزائري
    const cleanPhoneDigits = phone.replace(/[^0-9]/g, "");
    if (cleanPhoneDigits.length < 9) {
      return createJsonResponse({ success: false, error: "يرجى إدخال رقم هاتف جزائري صالح (مثال: 0797-30-96-43)." });
    }

    if (!city || city.length < 2 || city.length > 100) {
      return createJsonResponse({ success: false, error: "يرجى اختيار الولاية أو المدينة." });
    }

    if (!address || address.length < 3 || address.length > 250) {
      return createJsonResponse({ success: false, error: "يرجى كتابة العنوان الكامل للتوصيل." });
    }

    // التحقق من المقاس مع توفير قيمة افتراضية في حال عدم توفره
    if (!SERVER_CONFIG.ALLOWED_SIZES.includes(size)) {
      size = SERVER_CONFIG.ALLOWED_SIZES[0] || "40";
    }

    // التحقق من الكمية (بين 1 و 5)
    if (isNaN(quantity) || quantity < 1 || quantity > 5) {
      quantity = 1;
    }

    // 3. منع التكرار العرضي السريع (خلال 30 ثانية لنفس الهاتف والمقاس)
    const cache = CacheService.getScriptCache();
    const duplicateKey = "ord_" + cleanPhoneDigits + "_" + size;
    const existingOrderId = cache.get(duplicateKey);
    
    // 4. حساب الأسعار من السيرفر حصريًا
    const unitPrice = SERVER_CONFIG.UNIT_PRICE;
    const deliveryFee = (deliveryType === "yalidine") 
      ? SERVER_CONFIG.DELIVERY_FEES.yalidine 
      : SERVER_CONFIG.DELIVERY_FEES.domicile;
    const total = (unitPrice * quantity) + deliveryFee;

    // 5. الوصول إلى ورقة الطلبات
    const ss = getTargetSpreadsheet();
    let sheet = ss.getSheetByName(SERVER_CONFIG.SHEET_NAME);
    if (!sheet) {
      setupSheet();
      sheet = ss.getSheetByName(SERVER_CONFIG.SHEET_NAME);
    }

    // 6. توليد رقم الطلب التسلسلي الآمن TF-0001, TF-0002...
    const lastRow = sheet.getLastRow();
    let orderNumber = 1;
    
    if (lastRow > 1) {
      const lastIdValue = sheet.getRange(lastRow, 1).getValue().toString();
      const match = lastIdValue.match(/TF-(\d+)/);
      if (match) {
        orderNumber = parseInt(match[1], 10) + 1;
      } else {
        orderNumber = lastRow;
      }
    }
    
    const orderId = "TF-" + String(orderNumber).padStart(4, "0");

    // 7. توقيت الطلب بتوقيت الجزائر
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, SERVER_CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");

    // 8. إدراج اسم المنتج واللون معًا ليكون واضحًا مباشرة لصاحب المتجر
    const productFullLabel = SERVER_CONFIG.PRODUCT_NAME + " - " + color;

    // 9. إضافة سطر الطلب في الشيت بالترتيب المطلوب
    // Columns: Order ID | Date/Time | Status | Full Name | Phone | City | Address | Product | Size | Quantity | Unit Price | Delivery Fee | Total | Notes | My Notes
    const rowData = [
      orderId,
      formattedDate,
      SERVER_CONFIG.DEFAULT_STATUS,
      fullName,
      "'" + phone, // وضع علامة اقتباس لتثبيت الصفر الأولي في الهاتف
      city,
      address,
      productFullLabel,
      size,
      quantity,
      unitPrice,
      deliveryFee,
      total,
      notes,
      "" // خانة My Notes فارغة لملاحظات المتجر الشخصية
    ];

    sheet.appendRow(rowData);

    // تسجيل الطلب في الكاش لمدة 30 ثانية
    cache.put(duplicateKey, orderId, 30);

    return createJsonResponse({
      success: true,
      orderId: orderId,
      message: "تم استلام طلبك، سنتصل بك قريبًا لتأكيده"
    });

  } catch (error) {
    Logger.log("Error in handleOrderProcessing: " + error.toString());
    return createJsonResponse({
      success: false,
      error: "خطأ في معالجة الطلب: " + error.message
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * دالة تعقيم المدخلات لمنع حقن الصيغ والمعادلات في Google Sheets
 */
function sanitizeInput(str) {
  if (str === null || str === undefined) return "";
  let text = String(str).trim();
  if (/^[=\+\-@]/.test(text)) {
    text = "'" + text;
  }
  return text;
}

/**
 * دالة مساعدة لإرجاع استجابة JSON
 */
function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
