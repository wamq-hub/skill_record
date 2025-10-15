// ============================================
// ملف الإعدادات - config.js
// ============================================

const CONFIG = {

  // رابط Google Apps Script Web App
  // بعد نشر السكريبت، ضع الرابط هنا
  GOOGLE_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbwz0ttVWaZFN6FFLSdGVk6LjefPv5-jhT1kzOUhXdUrI80AN6F8Zv50mjzt3SpPGXUyLA/exec',
  
  // معرف Google Spreadsheet
  SPREADSHEET_ID: '1_pSdZAPagzI5qEbspfABCuGNlYy7TwzrS99TqfS8RMo',
  
  // أسماء الأوراق في Google Sheets
  SHEETS: {
    USERS: 'المستخدمين',           // ورقة بيانات المستخدمين
    WORKSHOPS: 'الورش',             // ورقة الورش المتاحة
    ATTENDANCE: 'الحضور',           // ورقة تسجيل الحضور
    SKILLS: 'المهارات',             // ورقة المهارات المكتسبة
    DEPARTMENTS: 'الأقسام',         // ورقة الأقسام
    WORKSHOP_TYPES: 'أنواع_الورش'   // ورقة أنواع الورش المعتمدة
  },
  

    EXTERNAL_CERTS: {
      MAX_FILE_SIZE: 5242880, // 5MB
      ALLOWED_EXTENSIONS: ['pdf', 'jpg', 'jpeg', 'png'],
      STATUS: {
        PENDING: 'بانتظار الاعتماد',
        APPROVED: 'معتمد', 
        REJECTED: 'مرفوض'
      }
    },
  // أنواع المستخدمين
  USER_TYPES: {
    TRAINEE: 'متدرب',
    TRAINER: 'مدرب',
    HEAD: 'رئيس_قسم',
    DEAN_STUDENTS: 'وكيل_شؤون_المتدربين',
    DEAN: 'عميد'
  },
  
  // حالات الحضور
  ATTENDANCE_STATUS: {
    PENDING: 'معلق',
    APPROVED: 'معتمد',
    REJECTED: 'مرفوض'
  },
  
  // حالات الورش
  WORKSHOP_STATUS: {
    AVAILABLE: 'متاح',
    FULL: 'مكتمل',
    COMPLETED: 'منتهي',
    CANCELLED: 'ملغي'
  },

  
  
  // قائمة المهارات المعتمدة (21 مهارة)
  SKILLS_LIST: [
    'مهارات العرض والإلقاء',
    'الإدارة المالية',
    'التفكير الإبداعي',
    'مهارات حل المشكلات',
    'مهارات الاتصال والتواصل',
    'المهارات القيادية',
    'التعامل مع ضغوط العمل',
    'بناء فريق العمل',
    'التخطيط الناجح',
    'إدارة الاجتماعات وإعداد محاضرها',
    'مهارات البحث عن وظائف',
    'مهارة اتخاذ القرار',
    'إدارة المعلومات',
    'اعداد وكتابة التقرير',
    'إدارة الوقت',
    'مهارات تطوير الذات',
    'اعداد السيرة الذاتية',
    'الريادة والمبادرة',
    'مهارات التفاوض',
    'السلامة المهنية',
    'العمل التطوعي'
  ],
  
  // إعدادات التطبيق
  APP_SETTINGS: {
    ENABLE_NOTIFICATIONS: true,
    AUTO_APPROVE: false,
    MAX_WORKSHOPS_PER_TRAINEE: 50,
    SESSION_TIMEOUT: 3600000 // 1 hour in milliseconds
  },
  

  // إعدادات التصدير من قالب Excel
  EXPORT: {
    TEMPLATE_URL: 'templates/skill_record.xlsx', // مسار ملف القالب داخل مشروعك
    SHEET_NAME:   'سجل المهارات',                // اسم ورقة القالب
    SKILLS_START_ROW: 9,                          // احتياطي/غير مستخدم عند وجود SKILLS_MAP
    TOTAL_HOURS_CELL: 'F30',                      // إجمالي الساعات في F30

    // رؤوس بيانات الطالب
    NAME_CELL:   'B5',
    ID_CELL:     'F5',
    MAJOR_CELL:  'B6',
    SEM_CELL:    'F6',
    YEAR_CELL:   'H6',

    // التواقيع
    HOD_CELL:      'B32', // رئيس القسم
    DEAN_STD_CELL: 'B33', // وكيل شؤون المتدربين
    DEAN_CELL:     'B34', // العميد

    // خريطة خلايا الساعات: تكتب الساعات في F9..F29
    // (الكود سيكتب اسم المهارة تلقائيًا في B{نفس الصف})
    SKILLS_MAP: {
      'مهارات العرض والإلقاء': 'F9',
      'الإدارة المالية': 'F10',
      'التفكير الإبداعي': 'F11',
      'مهارات حل المشكلات': 'F12',
      'مهارات الاتصال والتواصل': 'F13',
      'المهارات القيادية': 'F14',
      'التعامل مع ضغوط العمل': 'F15',
      'بناء فريق العمل': 'F16',
      'التخطيط الناجح': 'F17',
      'إدارة الاجتماعات وإعداد محاضرها': 'F18',
      'مهارات البحث عن وظائف': 'F19',
      'مهارة اتخاذ القرار': 'F20',
      'إدارة المعلومات': 'F21',
      'اعداد وكتابة التقرير': 'F22',
      'إدارة الوقت': 'F23',
      'مهارات تطوير الذات': 'F24',
      'اعداد السيرة الذاتية': 'F25',
      'الريادة والمبادرة': 'F26',
      'مهارات التفاوض': 'F27',
      'السلامة المهنية': 'F28',
      'العمل التطوعي': 'F29'
    }
  },


};


// دوال مساعدة للتحقق من نوع المستخدم
const UserHelpers = {
  isTrainee: (userType) => userType === CONFIG.USER_TYPES.TRAINEE,
  isTrainer: (userType) => userType === CONFIG.USER_TYPES.TRAINER,
  isHead: (userType) => userType === CONFIG.USER_TYPES.HEAD,
  isDeanStudents: (userType) => userType === CONFIG.USER_TYPES.DEAN_STUDENTS,
  isDean: (userType) => userType === CONFIG.USER_TYPES.DEAN,
  
  // التحقق من الصلاحيات الإدارية
  hasAdminAccess: (userType) => {
    return [
      CONFIG.USER_TYPES.HEAD,
      CONFIG.USER_TYPES.DEAN_STUDENTS,
      CONFIG.USER_TYPES.DEAN
    ].includes(userType);
  }
};
