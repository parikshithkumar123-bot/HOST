const APP_CONFIG = Object.freeze({
  appName: 'SAISETHSONS HOSPITALITY GROUP',
  scriptId: '17MoaU39_RAILhSWxihWW9D9ryldUGig4rtorY6olddviU1erNVmdRaw3',
  spreadsheetId: '1f0y1frkb4WPy4xvjKm6k_Iuq3EDnlfgwjJZE9Awb8es',
  sessionTimeoutMinutes: 30,
  sessionTtlSeconds: 1800,  // 30 minutes — must equal sessionTimeoutMinutes * 60
  cacheTtlSeconds: {
    boot: 300,
    dashboard: 180,
    hostLinks: 300,
    notifications: 60,
    newsletters: 120,
    hotzone: 180
  },
  sheets: Object.freeze({
    USERS: 'Users',
    RESTAURANTS: 'Restaurants',
    RESTAURANT_ACCESS: 'Restaurant Access',
    HOST_LINKS: 'HOST Links',
    NEWSLETTERS: 'Newsletters',
    BRANDS: 'Brands',
    ROLE_BUTTON_PERMISSIONS: 'Role Button Permissions',
    CONTACT_REQUESTS: 'Contact Requests',
    PASSWORD_CHANGE_REQUESTS: 'Password Change Requests',
    DROPDOWN_OPTIONS: 'Dropdown Options',
    REQUESTS_HR: 'Requests HR',
    REQUESTS_ADMIN: 'Requests Admin',
    REQUESTS_SUPPLIES: 'Requests Supplies',
    REQUESTS_MAINTENANCE: 'Requests Maintenance',
    NEW_HIRES: 'New Hires',
    TEAM_MEMBERS: 'Team Members',
    NOTIFICATIONS: 'Notifications',
    AUDIT_LOGS: 'Audit Logs',
    REPORTS_CACHE: 'Reports Cache',
    SETTINGS: 'Settings',
    USER_MANAGEMENT_RESPONSES: 'User Management Responses',
    OPERATIONS_RESPONSES: 'Operations Responses',
    EMPLOYEE_HISTORY: 'Employee History',
    HIRING_REQUESTS: 'Hiring Requests',
    ADMIN_REQUESTS: 'Admin Requests',
    PAYROLL_SUBMISSIONS: 'Payroll Submissions',
    PAYROLL_ROWS: 'Payroll Rows',
    PAYROLL_ROE_APPLICATIONS: 'Payroll ROE Applications',
    HOTZONE_EDITIONS: 'HotZone Editions',
    HOTZONE_CONTRIBUTIONS: 'HotZone Contributions'
  }),
  roles: Object.freeze({
    SUPER_ADMIN: 'SUPER_ADMIN',
    RESTAURANT_USER: 'RESTAURANT_USER',
    HR_ADMIN: 'HR_ADMIN',
    MRO: 'MRO',
    SENIOR_MANAGEMENT: 'SENIOR_MANAGEMENT'
  }),
  requestStatuses: Object.freeze([
    'PENDING',
    'APPROVED',
    'REJECTED',
    'IN_PROGRESS',
    'COMPLETED'
  ]),
  hireStatuses: Object.freeze([
    'APPLIED',
    'INTERVIEW_SCHEDULED',
    'SELECTED',
    'ORIENTATION_PENDING',
    'ORIENTATION_COMPLETE',
    'SENT_TO_RESTAURANT',
    'RESTAURANT_CONFIRMATION_PENDING',
    'STARTED_WORK',
    'ON_HOLD',
    'CANCELLED'
  ]),
  hostCategories: Object.freeze(['H', 'O', 'S', 'T'])
});

const ROLE_ACTIONS = Object.freeze({
  RECRUITMENT: 'RECRUITMENT',
  RESTAURANT_REQUEST: 'RESTAURANT_REQUEST',
  REPAIR_MAINTENANCE: 'REPAIR_MAINTENANCE',
  SNAPSHOT_5W: 'SNAPSHOT_5W',
  NEWSLETTER_MANAGEMENT: 'NEWSLETTER_MANAGEMENT',
  RECRUITMENT_REQUESTS: 'RECRUITMENT_REQUESTS',
  MAINTENANCE_REQUESTS: 'MAINTENANCE_REQUESTS',
  ASSIGNED_REPORTS: 'ASSIGNED_REPORTS',
  REPORTS: 'REPORTS',
  WEBSITE_CONTROL: 'WEBSITE_CONTROL',
  ROLE_ACTION_BUTTON_CONTROL: 'ROLE_ACTION_BUTTON_CONTROL',
  RESTAURANT_MANAGEMENT: 'RESTAURANT_MANAGEMENT',
  BRAND_MANAGEMENT: 'BRAND_MANAGEMENT',
  EMPLOYEE_MANAGEMENT: 'EMPLOYEE_MANAGEMENT',
  USER_MANAGEMENT: 'USER_MANAGEMENT',
  CONTACT_US_REQUESTS: 'CONTACT_US_REQUESTS',
  PASSWORD_CHANGE_REQUESTS: 'PASSWORD_CHANGE_REQUESTS',
  // FIX #6: DROPDOWN_SETTINGS is now a real routable action key.
  // It was previously only in HIDDEN_ROLE_ACTION_BUTTONS but missing from
  // ROLE_ACTIONS entirely, making the SECTION_DEF entry in ClientCore
  // unreachable. It stays hidden from the button control UI (see
  // HIDDEN_ROLE_ACTION_BUTTONS below) but can now be granted and routed.
  DROPDOWN_SETTINGS: 'DROPDOWN_SETTINGS',
  HOST_LINK_MANAGEMENT: 'HOST_LINK_MANAGEMENT',
  OPERATIONS_DASHBOARD: 'OPERATIONS_DASHBOARD',
  PAYROLL_SUBMISSION: 'PAYROLL_SUBMISSION',
  PAYROLL_REVIEW: 'PAYROLL_REVIEW',
  PAYROLL_ADMIN: 'PAYROLL_ADMIN',
  HOTZONE_MANAGEMENT: 'HOTZONE_MANAGEMENT',
  HOST_HOTZONE_VIEW: 'HOST_HOTZONE_VIEW',
  YEAR_END_TRANSITION: 'YEAR_END_TRANSITION'
});

const ALL_ROLE_ACTION_BUTTONS = Object.freeze([
  'ROLE_ACTION_BUTTON_CONTROL',
  'RESTAURANT_MANAGEMENT',
  'BRAND_MANAGEMENT',
  'EMPLOYEE_MANAGEMENT',
  'USER_MANAGEMENT',
  'RECRUITMENT',
  'RESTAURANT_REQUEST',
  'RECRUITMENT_REQUESTS',
  'CONTACT_US_REQUESTS',
  'PASSWORD_CHANGE_REQUESTS',
  'NEWSLETTER_MANAGEMENT',
  'HOST_LINK_MANAGEMENT',
  'OPERATIONS_DASHBOARD',
  'PAYROLL_SUBMISSION',
  'PAYROLL_REVIEW',
  'PAYROLL_ADMIN',
  'SNAPSHOT_5W',
  'HOTZONE_MANAGEMENT',
  'HOST_HOTZONE_VIEW',
  'YEAR_END_TRANSITION'
]);

// FIX #6: DROPDOWN_SETTINGS removed from HIDDEN_ROLE_ACTION_BUTTONS.
// It is still not shown in the role button control UI (it is not in
// ALL_ROLE_ACTION_BUTTONS) but is no longer incorrectly listed as
// a hidden button — that list is for buttons that exist but are hidden,
// not for routing keys that need to function independently.
const HIDDEN_ROLE_ACTION_BUTTONS = Object.freeze([
  'WEBSITE_CONTROL',
  'REPAIR_MAINTENANCE',
  'MAINTENANCE_REQUESTS',
  'ASSIGNED_REPORTS',
  'REPORTS'
]);

const ROLE_ACTION_NO_ACCESS_MARKER = '__NONE__';

const DEFAULT_ROLE_ACTION_PERMISSIONS = Object.freeze({
  SUPER_ADMIN: ALL_ROLE_ACTION_BUTTONS,

  RESTAURANT_USER: [
    'RECRUITMENT',
    'RESTAURANT_REQUEST',
    'PAYROLL_SUBMISSION',
    'SNAPSHOT_5W',
    'HOST_HOTZONE_VIEW'
  ],

  HR_ADMIN: [
    'HOTZONE_MANAGEMENT',
    'RECRUITMENT_REQUESTS',
    'RESTAURANT_REQUEST',
    'PAYROLL_REVIEW',
    'PAYROLL_ADMIN',
    'OPERATIONS_DASHBOARD'
  ],

  MRO: [
    'HOTZONE_MANAGEMENT',
    'RECRUITMENT_REQUESTS',
    'RESTAURANT_REQUEST',
    'PAYROLL_REVIEW',
    'PAYROLL_ADMIN',
    'OPERATIONS_DASHBOARD'
  ],

  SENIOR_MANAGEMENT: [
    'HOTZONE_MANAGEMENT',
    'RECRUITMENT_REQUESTS',
    'RESTAURANT_REQUEST',
    'PAYROLL_REVIEW',
    'PAYROLL_ADMIN',
    'OPERATIONS_DASHBOARD'
  ]
});

const BRAND = Object.freeze({
  companyName: 'SAISETHSONS HOSPITALITY GROUP',
  colors: Object.freeze({
    red: '#b51218',
    orange: '#e88d07',
    yellow: '#f5db07',
    cyan: '#2fc6d7',
    purple: '#8b2ee8',
    navy: '#163a63',
    charcoal: '#1d2430',
    surface: '#f7f9fc'
  })
});
