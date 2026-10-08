/**
 * =========================================================================
 * TAMARA FASHION - Google Apps Script Backend (Code.gs)
 * =========================================================================
 * هذا السكربت يقوم بحفظ طلبات متجر TAMARA FASHION تلقائيًا في Google Sheet
 * مع الحماية الكاملة من السبام وتكرار الطلبات وتأمين الأسعار وحقن المعادلات.
 * 
 * المطور لـ: TAMARA FASHION
 * المنتوج: فستان ميرال قطيفة طقم كامل 4 قطع
 * =========================================================================
 */

// إعدادات المتجر على السيرفر (لا يمكن للعميل أو المتصفح التلاعب بها)
const SERVER_CONFIG = {
  SHEET_NAME: "Orders",
  PRODUCT_NAME: "فستان ميرال قطيفة طقم كامل 4 قطع",
  UNIT_PRICE: 4900, // السعر الثابت 4900 دج
  DELIVERY_FEES: {
    domicile: 700, // توصيل للمنزل: 700 دج
    yalidine: 550  // استلام من مكتب ياليدين: 550 دج
  },
  ALLOWED_SIZES: ["36", "38", "40", "42", "44", "46", "48", "50"],
  STATUS_OPTIONS: ["جديد", "مؤكد", "قيد التحضير", "تم الشحن", "تم التسليم", "ملغي"],
  DEFAULT_STATUS: "جديد",
  TIMEZONE: "Africa/Algiers"
};

/**
 * دالة الإعداد التلقائي بضغطة زر واحدة (Run setupSheet)
 * تقوم بتهيئة الورقة "Orders"، إضافة العناوين، تجميد السطر الأول، 
 * وتعيين القوائم المنسدلة لحالة الطلب مع ألوان التمييز التلقائية.
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SERVER_CONFIG.SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SERVER_CONFIG.SHEET_NAME);
  }
  
  // عناوين الأعمدة المطلوبة
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
 * فحص الاتصال البسيط (GET Request) للتأكد من تشغيل الـ Web App
 */
function doGet(e) {
  return createJsonResponse({
    status: "ok",
    message: "TAMARA FASHION Order API is running."
  });
}

/**
 * استقبال ومعالجة طلبات الشراء الجديدة (POST Request)
 */
function doPost(e) {
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
    
    // استخراج البيانات سواء تم إرسالها بتنسيق JSON أو Form
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e.parameter) {
      data = e.parameter;
    }

    // 1. فحص مصيدة الروبوتات (Honeypot): إذا امتلأ الحقل المخفي يتم تجاهل الطلب
    if (data.website || data.honeypot || data.middle_name) {
      // إرجاع نجاح وهمي حتى لا يعيد البوت المحاولة
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
    const size = String(data.size || "").trim();
    let quantity = parseInt(data.quantity, 10);
    const deliveryType = String(data.deliveryType || "domicile").toLowerCase().trim();
    const notes = sanitizeInput(data.notes || "");

    if (!fullName || fullName.length < 2 || fullName.length > 100) {
      return createJsonResponse({ success: false, error: "يرجى كتابة الاسم الكامل بشكل صحيح." });
    }

    // التحقق من صحة رقم الهاتف الجزائري (05 / 06 / 07 متبوع بـ 8 أرقام، أو +213)
    const cleanPhoneDigits = phone.replace(/[^0-9]/g, "");
    const phoneRegex = /^(0|\+?213)(5|6|7)[0-9]{8}$/;
    if (!phoneRegex.test(phone.replace(/\s+/g, "")) && cleanPhoneDigits.length < 9) {
      return createJsonResponse({ success: false, error: "يرجى إدخال رقم هاتف جزائري صالح (مثال: 0797309643)." });
    }

    if (!city || city.length < 2 || city.length > 100) {
      return createJsonResponse({ success: false, error: "يرجى اختيار الولاية أو المدينة." });
    }

    if (!address || address.length < 3 || address.length > 250) {
      return createJsonResponse({ success: false, error: "يرجى كتابة العنوان الكامل للتوصيل." });
    }

    // التحقق من المقاس
    if (!SERVER_CONFIG.ALLOWED_SIZES.includes(size)) {
      return createJsonResponse({ success: false, error: "يرجى اختيار مقاس صالح من المقاسات المتوفرة." });
    }

    // التحقق من الكمية (بين 1 و 5)
    if (isNaN(quantity) || quantity < 1 || quantity > 5) {
      quantity = 1;
    }

    // 3. منع الطلبات المكررة (خلال دقيقتين لنفس الهاتف ونفس المقاس)
    const cache = CacheService.getScriptCache();
    const duplicateKey = "ord_" + cleanPhoneDigits + "_" + size;
    const existingOrderId = cache.get(duplicateKey);
    
    if (existingOrderId) {
      // إرجاع نفس رقم الطلب السابق بدون تكرار الحفظ في الشيت
      return createJsonResponse({
        success: true,
        orderId: existingOrderId,
        message: "تم استلام طلبك، سنتصل بك قريبًا لتأكيده (طلب مسجل مسبقًا)"
      });
    }

    // 4. حساب الأسعار من السيرفر حصريًا (عدم الاعتماد على أي سعر قادم من المتصفح)
    const unitPrice = SERVER_CONFIG.UNIT_PRICE;
    const deliveryFee = (deliveryType === "yalidine") 
      ? SERVER_CONFIG.DELIVERY_FEES.yalidine 
      : SERVER_CONFIG.DELIVERY_FEES.domicile;
    const total = (unitPrice * quantity) + deliveryFee;

    // 5. الوصول إلى ورقة الطلبات
    const ss = SpreadsheetApp.getActiveSpreadsheet();
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
        orderNumber = lastRow; // احتياطي
      }
    }
    
    const orderId = "TF-" + String(orderNumber).padStart(4, "0");

    // 7. توقيت الطلب بتوقيت الجزائر
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, SERVER_CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");

    // 8. إضافة سطر الطلب في الشيت بالترتيب المطلوب
    // Columns: Order ID | Date/Time | Status | Full Name | Phone | City | Address | Product | Size | Quantity | Unit Price | Delivery Fee | Total | Notes | My Notes
    const rowData = [
      orderId,
      formattedDate,
      SERVER_CONFIG.DEFAULT_STATUS,
      fullName,
      "'" + phone, // وضع علامة اقتباس لتثبيت الصفر الأولي في الهاتف
      city,
      address,
      SERVER_CONFIG.PRODUCT_NAME,
      size,
      quantity,
      unitPrice,
      deliveryFee,
      total,
      notes,
      "" // خانة My Notes فارغة لملاحظات المتجر الشخصية
    ];

    sheet.appendRow(rowData);

    // تسجيل الطلب في الكاش لمدة دقيقتين (120 ثانية) لمنع التكرار
    cache.put(duplicateKey, orderId, 120);

    return createJsonResponse({
      success: true,
      orderId: orderId,
      message: "تم استلام طلبك، سنتصل بك قريبًا لتأكيده"
    });

  } catch (error) {
    Logger.log("Error in doPost: " + error.toString());
    // عدم كشف أخطاء تقنية للزبون
    return createJsonResponse({
      success: false,
      error: "عذرًا، حدث خطأ أثناء إرسال الطلب. يرجى المحاولة مرة أخرى أو الاتصال بنا مباشرة."
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * دالة تعقيم المدخلات لمنع حقن الصيغ والمعادلات في Google Sheets
 * Formula Injection Protection:
 * إذا كان النص يبدأ بـ = أو + أو - أو @ يضاف قبله رمز ' حتى يعامل كنص خالص.
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
 * دالة مساعدة لإرجاع استجابة JSON مع الترويسات السليمة للمتصفح
 */
function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
