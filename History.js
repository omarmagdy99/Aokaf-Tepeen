// // ========================================
// // نظام إدارة خطباء التبين - بدون CORS Issues
// // ========================================

// const CONFIG = {
//   PREACHERS_SHEET: "أسماء الخطباء",
//   SERMONS_SHEET: "الخطب",
//   MAYO_SHEET: "خطب 15 مايو"
// };

// // ========================================
// // doGet - لقراءة وجلب البيانات
// // ========================================
// function doGet(e) {
//   // ✅ التحقق: هل فيه parameter اسمه "action"؟
//   if (e && e.parameter && e.parameter.action === "updateStatus") {
//     // طلب حفظ حالة
//     return handleUpdateStatus(e.parameter);
//   }
  
//   // طلب جلب البيانات العادي
//   return handleGetData();
// }

// // ========================================
// // جلب البيانات (قراءة)
// // ========================================
// function handleGetData() {
//   try {
//     const preachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.PREACHERS_SHEET);
//     const preachersData = preachersSheet ? preachersSheet.getRange("A:F").getValues() : [];
    
//     const sermonsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SERMONS_SHEET);
//     const sermonsData = sermonsSheet ? sermonsSheet.getRange("A:F").getValues() : [];
    
//     const mayoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MAYO_SHEET);
//     let mayoData = mayoSheet ? mayoSheet.getRange("A:C").getValues() : [];
    
//     // بناء الحالات المحفوظة
//     const preachersStatus = {};
//     for (let i = 1; i < preachersData.length; i++) {
//       const row = preachersData[i];
//       if (row && row[1] && row[1].toString().trim() !== "") {
//         const nationalId = row[1].toString().trim();
//         preachersStatus[nationalId] = {
//           name: row[0]?.toString().trim() || "",
//           nationalId: nationalId,
//           statuses: {
//             1: row[2]?.toString().trim() || "",
//             2: row[3]?.toString().trim() || "",
//             3: row[4]?.toString().trim() || "",
//             4: row[5]?.toString().trim() || ""
//           }
//         };
//       }
//     }
    
//     return ContentService
//       .createTextOutput(JSON.stringify({
//         success: true,
//         preachers: preachersData,
//         sermons: sermonsData,
//         mayo: mayoData,
//         preachersStatus: preachersStatus
//       }))
//       .setMimeType(ContentService.MimeType.JSON);
      
//   } catch(error) {
//     return ContentService
//       .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
//       .setMimeType(ContentService.MimeType.JSON);
//   }
// }

// // ========================================
// // تحديث الحالة عبر GET parameters
// // ========================================
// function handleUpdateStatus(params) {
//   try {
//     const nationalId = params.nationalId;
//     const sermonOrder = parseInt(params.sermonOrder);
//     const status = params.status;
    
//     if (!nationalId) {
//       return ContentService
//         .createTextOutput(JSON.stringify({ success: false, error: "الرقم السري مطلوب" }))
//         .setMimeType(ContentService.MimeType.JSON);
//     }
    
//     const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.PREACHERS_SHEET);
//     if (!sheet) {
//       return ContentService
//         .createTextOutput(JSON.stringify({ success: false, error: "الشيت غير موجود" }))
//         .setMimeType(ContentService.MimeType.JSON);
//     }
    
//     const allData = sheet.getRange("A:F").getValues();
    
//     let rowIndex = -1;
//     for (let i = 1; i < allData.length; i++) {
//       const cellValue = allData[i][1]?.toString().trim();
//       if (cellValue === nationalId) {
//         rowIndex = i + 1;
//         break;
//       }
//     }
    
//     if (rowIndex === -1) {
//       return ContentService
//         .createTextOutput(JSON.stringify({ success: false, error: "الرقم السري غير مسجل" }))
//         .setMimeType(ContentService.MimeType.JSON);
//     }
    
//     const columnIndex = 2 + sermonOrder;
//     sheet.getRange(rowIndex, columnIndex).setValue(status);
    
//     return ContentService
//       .createTextOutput(JSON.stringify({ success: true, message: "تم الحفظ" }))
//       .setMimeType(ContentService.MimeType.JSON);
      
//   } catch(error) {
//     return ContentService
//       .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
//       .setMimeType(ContentService.MimeType.JSON);
//   }
// }



// ========================================
// نظام إدارة خطباء التبين - النسخة المطورة
// ========================================

const CONFIG = {
  PREACHERS_SHEET: "أسماء الخطباء",
  SERMONS_SHEET: "الخطب",
  MAYO_SHEET: "خطب 15 مايو",
  ADMINS_SHEET: "المشرفين"  // ✅ شيت جديد
};

// ========================================
// doGet - المدخل الرئيسي
// ========================================
function doGet(e) {
  // جلب الـ parameter page
  const page = e?.parameter?.page || "";
  
  // 👑 صفحة اللوجين (بتنفع تتعرض في iframe أو مباشرة)
  if (page === "admin") {
    return HtmlService.createHtmlOutputFromFile("admin")
      .setTitle("لوحة تحكم المشرفين")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // 🔐 طرق API الخاصة بالمشرفين
  if (e && e.parameter) {
    const action = e.parameter.action;
    
    // تسجيل الدخول
    if (action === "adminLogin") {
      return handleAdminLogin(e.parameter);
    }
    
    // جلب كل المشرفين
    if (action === "getAdmins") {
      return handleGetAdmins();
    }
    
    // إضافة مشرف جديد
    if (action === "addAdmin") {
      return handleAddAdmin(e.parameter);
    }
    
    // تحديث مشرف
    if (action === "updateAdmin") {
      return handleUpdateAdmin(e.parameter);
    }
    
    // حذف مشرف
    if (action === "deleteAdmin") {
      return handleDeleteAdmin(e.parameter);
    }
    
    // جلب إحصائيات الـ Dashboard
    if (action === "getDashboardStats") {
      return handleGetDashboardStats(e.parameter);
    }
    
    // تحديث حالة خطيب
    if (action === "updateStatus") {
      return handleUpdateStatus(e.parameter);
    }
  }
  
  // الصفحة العادية للخطباء
  return handleGetData();
}

// ========================================
// 🔐 التعامل مع المشرفين (CRUD)
// ========================================

// ✅ تسجيل الدخول
function handleAdminLogin(params) {
  try {
    const username = params.username?.toString().trim();
    const password = params.password?.toString().trim();
    
    if (!username || !password) {
      return createResponse(false, "اسم المستخدم وكلمة المرور مطلوبة");
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ADMINS_SHEET);
    if (!sheet) {
      return createResponse(false, "شيت المشرفين غير موجود");
    }
    
    const data = sheet.getDataRange().getValues();
    
    // تخطي رأس الجدول
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]?.toString().trim() === username && row[1]?.toString().trim() === password) {
        return createResponse(true, "تم تسجيل الدخول بنجاح", {
          username: username,
          role: row[2]?.toString().trim() || "admin"
        });
      }
    }
    
    return createResponse(false, "اسم المستخدم أو كلمة المرور غير صحيحة");
    
  } catch (error) {
    return createResponse(false, error.toString());
  }
}

// ✅ جلب كل المشرفين
function handleGetAdmins() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ADMINS_SHEET);
    if (!sheet) {
      return createResponse(false, "شيت المشرفين غير موجود");
    }
    
    const data = sheet.getDataRange().getValues();
    const admins = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[0].toString().trim() !== "") {
        admins.push({
          username: row[0].toString().trim(),
          password: row[1]?.toString().trim() || "",
          role: row[2]?.toString().trim() || "admin",
          createdAt: row[3]?.toString().trim() || ""
        });
      }
    }
    
    return createResponse(true, "تم جلب البيانات", { admins: admins });
    
  } catch (error) {
    return createResponse(false, error.toString());
  }
}

// ✅ إضافة مشرف جديد
function handleAddAdmin(params) {
  try {
    const username = params.username?.toString().trim();
    const password = params.password?.toString().trim();
    const role = params.role?.toString().trim() || "admin";
    
    if (!username || !password) {
      return createResponse(false, "اسم المستخدم وكلمة المرور مطلوبة");
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ADMINS_SHEET);
    if (!sheet) {
      return createResponse(false, "شيت المشرفين غير موجود");
    }
    
    // التحقق من عدم وجود نفس المستخدم
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]?.toString().trim() === username) {
        return createResponse(false, "اسم المستخدم موجود بالفعل");
      }
    }
    
    const createdAt = new Date().toLocaleString("ar-EG");
    sheet.appendRow([username, password, role, createdAt]);
    
    return createResponse(true, "تم إضافة المشرف بنجاح");
    
  } catch (error) {
    return createResponse(false, error.toString());
  }
}

// ✅ تحديث مشرف
function handleUpdateAdmin(params) {
  try {
    const oldUsername = params.oldUsername?.toString().trim();
    const newUsername = params.newUsername?.toString().trim();
    const newPassword = params.newPassword?.toString().trim();
    const newRole = params.newRole?.toString().trim();
    
    if (!oldUsername || !newUsername || !newPassword) {
      return createResponse(false, "جميع الحقول مطلوبة");
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ADMINS_SHEET);
    if (!sheet) {
      return createResponse(false, "شيت المشرفين غير موجود");
    }
    
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]?.toString().trim() === oldUsername) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return createResponse(false, "المشرف غير موجود");
    }
    
    // تحديث البيانات
    sheet.getRange(rowIndex, 1).setValue(newUsername);
    sheet.getRange(rowIndex, 2).setValue(newPassword);
    if (newRole) sheet.getRange(rowIndex, 3).setValue(newRole);
    
    return createResponse(true, "تم تحديث المشرف بنجاح");
    
  } catch (error) {
    return createResponse(false, error.toString());
  }
}

// ✅ حذف مشرف
function handleDeleteAdmin(params) {
  try {
    const username = params.username?.toString().trim();
    
    if (!username) {
      return createResponse(false, "اسم المستخدم مطلوب");
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ADMINS_SHEET);
    if (!sheet) {
      return createResponse(false, "شيت المشرفين غير موجود");
    }
    
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]?.toString().trim() === username) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return createResponse(false, "المشرف غير موجود");
    }
    
    sheet.deleteRow(rowIndex);
    
    return createResponse(true, "تم حذف المشرف بنجاح");
    
  } catch (error) {
    return createResponse(false, error.toString());
  }
}

// ========================================
// 📊 إحصائيات الـ Dashboard
// ========================================
function handleGetDashboardStats(params) {
  try {
    const selectedSermon = params?.sermon || 1; // الخطبة المختارة (1-4)
    
    // جلب بيانات الخطباء
    const preachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.PREACHERS_SHEET);
    const preachersData = preachersSheet ? preachersSheet.getRange("A:F").getValues() : [];
    
    // جلب بيانات الخطب
    const sermonsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SERMONS_SHEET);
    const sermonsData = sermonsSheet ? sermonsSheet.getRange("A:F").getValues() : [];
    
    // إحصائيات الخطباء
    let confirmedCount = 0;      // مؤكدين
    let apologizedCount = 0;     // معتذرين
    let readyCount = 0;          // مستعدين للخطابة
    
    // الخطباء المستعدين (حالتهم "مستعد للخطبة" في أي خطبة)
    const readyPreachers = [];
    
    // المساجد المحتاجة
    const needyMosques = [];
    
    // 1️⃣ تحليل حالات الخطباء
    for (let i = 1; i < preachersData.length; i++) {
      const row = preachersData[i];
      if (!row || !row[1]) continue;
      
      const status1 = row[2]?.toString().trim() || "";
      const status2 = row[3]?.toString().trim() || "";
      const status3 = row[4]?.toString().trim() || "";
      const status4 = row[5]?.toString().trim() || "";
      
      // للخطبة المحددة
      const selectedStatus = [status1, status2, status3, status4][selectedSermon - 1];
      if (selectedStatus === "تأكيد الحضور") confirmedCount++;
      if (selectedStatus === "معتذر") apologizedCount++;
      
      // المستعدين للخطابة (في أي خطبة)
      if (status1 === "مستعد للخطبة" || status2 === "مستعد للخطبة" || 
          status3 === "مستعد للخطبة" || status4 === "مستعد للخطبة") {
        readyCount++;
        readyPreachers.push(row[0]?.toString().trim());
      }
    }
    
    // 2️⃣ تحليل المساجد المحتاجة
    if (sermonsData.length > 1) {
      for (let i = 1; i < sermonsData.length; i++) {
        const row = sermonsData[i];
        if (!row || row.length < 6) continue;
        
        const mosqueName = row[0]?.toString().trim() || "";
        const area = row[1]?.toString().trim() || "";
        const preacher = row[selectedSermon + 1]?.toString().trim() || ""; // +1 عشان العمود A=index0
        
        // نبحث عن حالة هذا الخطيب في الخطبة المحددة
        let preacherStatus = "";
        for (let j = 1; j < preachersData.length; j++) {
          const pRow = preachersData[j];
          if (pRow && pRow[0]?.toString().trim() === preacher) {
            const statuses = [pRow[2], pRow[3], pRow[4], pRow[5]];
            preacherStatus = statuses[selectedSermon - 1]?.toString().trim() || "";
            break;
          }
        }
        
        // المسجد محتاج لو: مفيش خطيب OR الخطيب معتذر
        if (!preacher || preacher === "" || preacherStatus === "معتذر") {
          needyMosques.push({
            mosque: area ? `${mosqueName} - ${area}` : mosqueName,
            currentPreacher: preacher || "لا يوجد",
            status: preacherStatus === "معتذر" ? "معتذر" : "ليس هناك خطيب"
          });
        }
      }
    }
    
    // 3️⃣ الحصول على أسماء الخطبة
    let sermonNames = ["الخطبة الأولى", "الخطبة الثانية", "الخطبة الثالثة", "الخطبة الرابعة"];
    if (sermonsData[0] && sermonsData[0].length >= 6) {
      sermonNames = [
        sermonsData[0][2]?.toString().trim() || sermonNames[0],
        sermonsData[0][3]?.toString().trim() || sermonNames[1],
        sermonsData[0][4]?.toString().trim() || sermonNames[2],
        sermonsData[0][5]?.toString().trim() || sermonNames[3]
      ];
    }
    
    return createResponse(true, "تم جلب الإحصائيات", {
      stats: {
        confirmed: confirmedCount,
        apologized: apologizedCount,
        ready: readyCount,
        needyMosques: needyMosques.length,
        totalMosques: sermonsData.length - 1
      },
      sermonNames: sermonNames,
      needyMosques: needyMosques,
      readyPreachers: readyPreachers
    });
    
  } catch (error) {
    return createResponse(false, error.toString());
  }
}

// ========================================
// 🛠️ دالة مساعدة للردود
// ========================================
function createResponse(success, message, data = null) {
  const response = { success: success, message: message };
  if (data) response.data = data;
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========================================
// 📝 باقي الدوال القديمة (handleGetData, handleUpdateStatus)
// ========================================

// جلب البيانات (قراءة) - موجودة قبل كده
function handleGetData() {
  try {
    const preachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.PREACHERS_SHEET);
    const preachersData = preachersSheet ? preachersSheet.getRange("A:F").getValues() : [];
    
    const sermonsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SERMONS_SHEET);
    const sermonsData = sermonsSheet ? sermonsSheet.getRange("A:F").getValues() : [];
    
    const mayoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MAYO_SHEET);
    let mayoData = mayoSheet ? mayoSheet.getRange("A:C").getValues() : [];
    
    // بناء الحالات المحفوظة
    const preachersStatus = {};
    for (let i = 1; i < preachersData.length; i++) {
      const row = preachersData[i];
      if (row && row[1] && row[1].toString().trim() !== "") {
        const nationalId = row[1].toString().trim();
        preachersStatus[nationalId] = {
          name: row[0]?.toString().trim() || "",
          nationalId: nationalId,
          statuses: {
            1: row[2]?.toString().trim() || "",
            2: row[3]?.toString().trim() || "",
            3: row[4]?.toString().trim() || "",
            4: row[5]?.toString().trim() || ""
          }
        };
      }
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        preachers: preachersData,
        sermons: sermonsData,
        mayo: mayoData,
        preachersStatus: preachersStatus
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// تحديث الحالة - موجودة قبل كده
function handleUpdateStatus(params) {
  try {
    const nationalId = params.nationalId;
    const sermonOrder = parseInt(params.sermonOrder);
    const status = params.status;
    
    if (!nationalId) {
      return createResponse(false, "الرقم السري مطلوب");
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.PREACHERS_SHEET);
    if (!sheet) {
      return createResponse(false, "الشيت غير موجود");
    }
    
    const allData = sheet.getRange("A:F").getValues();
    
    let rowIndex = -1;
    for (let i = 1; i < allData.length; i++) {
      const cellValue = allData[i][1]?.toString().trim();
      if (cellValue === nationalId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return createResponse(false, "الرقم السري غير مسجل");
    }
    
    const columnIndex = 2 + sermonOrder;
    sheet.getRange(rowIndex, columnIndex).setValue(status);
    
    return createResponse(true, "تم الحفظ");
    
  } catch(error) {
    return createResponse(false, error.toString());
  }
}