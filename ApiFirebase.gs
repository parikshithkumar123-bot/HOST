function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var action = body.action;
    var data = body.data || {};

    var result;

    if (action === 'login') {
      result = login(data);
      result = firebaseRememberSession_(result);

    } else if (action === 'logout') {
      firebaseForgetSession_(data.token);
      result = safeCallOne_([
        ['logout', [data.token]],
        ['logoutUser', [data.token]]
      ], { success: true, data: {} });

    } else if (action === 'dashboardBoot') {
      result = firebaseDashboardBoot_(firebaseResolveToken_(data.token));

    } else if (action === 'adminWorkspaceLoad') {
      var awToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['getAdminWorkspacePayload', [awToken]],
        ['getAdminToolsPayload', [awToken]]
      ], { success: true, data: { adminTools: {}, roleActions: [] } });

    } else if (action === 'manageTeamLoad') {
      var token = firebaseResolveToken_(data.token);
      result = getManageTeamData(token);
      if (result && result.success && result.data) {
        result.data.restaurants = result.data.restaurants || [];
        result.data.teamMembers = result.data.teamMembers || result.data.members || result.data.employees || [];
        result.data.members = result.data.teamMembers;
        result.data.employees = result.data.teamMembers;
        result.data.transferHistory = result.data.transferHistory || [];
      }

    } else if (action === 'manageTeamSaveEmployee') {
      var saveToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['saveTeamMemberFromManageTeam', [{ sessionToken: saveToken, newlyHired: data.newlyHired, status: data.status, name: data.name, mobile: data.mobile, hiredDate: data.hiredDate, assignedRestaurants: data.assignedRestaurants || [] }]],
        ['saveManageTeamEmployee', [data]],
        ['createManageTeamEmployee', [data]]
      ], null);

    } else if (action === 'manageTeamUpdateEmployee') {
      var updateToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['updateManageTeamEmployeeDetails', [{ sessionToken: updateToken, teamMemberId: data.teamMemberId, name: data.name, mobile: data.mobile, hiredDate: data.hiredDate, assignedRestaurants: data.assignedRestaurants || [] }]],
        ['saveManageTeamEmployeeDetails', [data]]
      ], null);
      if (result && result.success && data.status) {
        safeCallOne_([
          ['updateManageTeamEmployeeStatus', [{ sessionToken: updateToken, teamMemberId: data.teamMemberId, status: data.status }]],
          ['updateNewHeroRecruitmentStatus', [{ sessionToken: updateToken, id: data.teamMemberId, status: data.status, quitReason: '' }]]
        ], { success: true, data: {} });
      }

    } else if (action === 'manageTeamDeactivateEmployee') {
      var deactivateToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['deactivateManageTeamEmployee', [{ sessionToken: deactivateToken, teamMemberId: data.teamMemberId }]],
        ['updateManageTeamEmployeeStatus', [{ sessionToken: deactivateToken, teamMemberId: data.teamMemberId, status: 'INACTIVE' }]]
      ], null);

    } else if (action === 'manageTeamTransferEmployee') {
      var transferToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['transferManageTeamEmployee', [{ sessionToken: transferToken, teamMemberId: data.teamMemberId, toRestaurantId: data.toRestaurantId, transferDate: data.transferDate, notes: data.notes || '' }]]
      ], null);

    } else if (action === 'payrollSubmitSave') {
      var submitToken = firebaseResolveToken_(data.token);
      var payload = data.payload || {};
      payload.sessionToken = submitToken;
      result = safeCallOne_([
        ['submitPayrollV360', [payload]],
        ['savePayrollSubmissionV360', [payload]],
        ['submitPayrollSubmission', [payload]],
        ['submitPayrollOnBehalf', [payload]],
        ['savePayrollSubmission', [payload]]
      ], { success: true, data: { message: 'Captured.' } });

    } else if (action === 'payrollSubmissionLoad') {
      var payrollToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['getPayrollSubmissionWorkspaceData', [payrollToken]],
        ['getPayrollWorkspaceData', [payrollToken]]
      ], { success: true, data: { restaurants: [], submissions: [], roeApplications: [] } });
      result = firebaseNormalizePayroll_(result);

    } else if (action === 'payrollReviewLoad') {
      var payrollReviewToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['getPayrollReviewWorkspaceData', [payrollReviewToken]],
        ['getPayrollReviewData', [payrollReviewToken]]
      ], { success: true, data: { restaurants: [], submissions: [] } });
      result = firebaseNormalizePayroll_(result);

    } else if (action === 'payrollAdminLoad') {
      var payrollAdminToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['queryPayrollAdminWorkspaceData', [{ sessionToken: payrollAdminToken }]],
        ['getPayrollAdminWorkspaceData', [payrollAdminToken]]
      ], { success: true, data: { restaurants: [], approvedPayrolls: [] } });
      result = firebaseNormalizePayroll_(result);

    } else if (action === 'payrollMRODash') {
      var pmroToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['getMROPayrollDashboardData', [pmroToken]]
      ], { success: true, data: { pendingPayrolls: [], restaurants: [] } });

    } else if (action === 'payrollRoeLoad') {
      var prlToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['getPayrollRoeApplicationData', [{ sessionToken: prlToken }]]
      ], { success: true, data: { roeApplications: [] } });

    } else if (action === 'payrollStatusUpdate') {
      var psuToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['updatePayrollSubmissionStatus', [{ sessionToken: psuToken, submissionId: data.submissionId, status: data.status, adminNotes: data.adminNotes || '', correctionNotes: data.correctionNotes || '' }]]
      ], null);

    } else if (action === 'payrollRoeStatusUpdate') {
      var prsuToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['updatePayrollRoeStatus', [{ sessionToken: prsuToken, roeId: data.roeId, status: data.status, notes: data.notes || '' }]]
      ], null);

    } else if (action === 'payrollAdminWebsiteUpdate') {
      var pawToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['updatePayrollAdminWebsiteStatus', [Object.assign({ sessionToken: pawToken }, data)]]
      ], null);

    } else if (action === 'payrollRoeApply') {
      var praToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['submitPayrollRoeApplication', [Object.assign({ sessionToken: praToken }, data)]]
      ], null);

    } else if (action === 'payrollRecall') {
      var precToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['recallPayrollSubmission', [{ sessionToken: precToken, submissionId: data.submissionId }]]
      ], null);

    } else if (action === 'payrollReturn') {
      var pretToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['returnPayrollSubmissionToRestaurant', [{ sessionToken: pretToken, submissionId: data.submissionId }]]
      ], null);

    } else if (action === 'payrollExportCsv') {
      var pecToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['exportPayrollAdminCsv', [{ sessionToken: pecToken }]]
      ], { success: true, data: { csvUrl: '' } });

    } else if (action === 'payrollExportSummary') {
      var pesToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['exportPayrollSummary', [{ sessionToken: pesToken }]]
      ], { success: true, data: { csvUrl: '' } });

    } else if (action === 'payrollDeactivateRoe') {
      var pdrToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['deactivateEmployeeFromRoe', [{ sessionToken: pdrToken, roeId: data.roeId }]]
      ], null);

    } else if (action === 'recruitmentLoad') {
      var recruitmentToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['getRecruitmentData', [recruitmentToken]],
        ['getNewHeroEmployees', [recruitmentToken]],
        ['getManageTeamData', [recruitmentToken]]
      ], { success: true, data: { restaurants: [], teamMembers: [], hiringRequests: [] } });
      result = firebaseNormalizeRecruitment_(result);

    } else if (action === 'recruitmentStatusUpdate') {
      var rsuToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['updateNewHeroRecruitmentStatus', [Object.assign({ sessionToken: rsuToken }, data)]],
        ['updateNewHeroEmployeeStatus', [Object.assign({ sessionToken: rsuToken }, data)]]
      ], null);

    } else if (action === 'recruitmentSaveEmployee') {
      var rseToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['saveNewHeroEmployee', [Object.assign({ sessionToken: rseToken }, data)]]
      ], null);

    } else if (action === 'recruitmentCreateHiring') {
      var rchToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['createHiringRequest', [Object.assign({ sessionToken: rchToken }, data)]]
      ], null);

    } else if (action === 'restaurantRequestsLoad') {
      var rrToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['getRestaurantRequestReceiverData', [rrToken]],
        ['getRestaurantRequestsData', [rrToken]]
      ], { success: true, data: { restaurants: [], hiringRequests: [], adminRequests: [] } });

    } else if (action === 'restaurantRequestStatusUpdate') {
      var rrsuToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['updateRestaurantRequestReceiverStatus', [{ sessionToken: rrsuToken, requestId: data.requestId, status: data.status }]]
      ], null);

    } else if (action === 'adminRequestsLoad') {
      var arToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['getAdminRequestWorkspaceData', [arToken]],
        ['getAdminRequestsData', [arToken]]
      ], { success: true, data: { requests: [], restaurants: [] } });

    } else if (action === 'adminRequestStatusUpdate') {
      var arsuToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['updateAdminRequestStatus', [{ sessionToken: arsuToken, requestId: data.requestId, status: data.status, notes: data.notes || '' }]]
      ], null);

    } else if (action === 'adminRequestCreate') {
      var arcToken = firebaseResolveToken_(data.token);
      var arcPayload = Object.assign({ sessionToken: arcToken }, data);
      if (data.type === 'JOB_LETTER') {
        result = safeCallOne_([['createAdminJobLetterRequest', [arcPayload]]], null);
      } else {
        result = safeCallOne_([['createAdminOfficeSupplyRequest', [arcPayload]]], null);
      }

    } else if (action === 'changeOwnPassword') {
      var copToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([
        ['changeOwnPortalPassword', [Object.assign({ sessionToken: copToken }, data)]]
      ], null);

    } else if (action === 'contactRequestsLoad') {
      var crlToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['listContactRequests', [crlToken]]], { success: true, data: { requests: [] } });

    } else if (action === 'contactRequestCreate') {
      var crcToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['submitContactRequest', [Object.assign({ sessionToken: crcToken }, data)]]], null);

    } else if (action === 'contactRequestUpdate') {
      var cruToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['updateContactRequest', [Object.assign({ sessionToken: cruToken }, data)]]], null);

    } else if (action === 'passwordChangeRequestsLoad') {
      var pcrlToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['listPasswordChangeRequests', [pcrlToken]]], { success: true, data: { requests: [] } });

    } else if (action === 'passwordChangeRequestCreate') {
      var pcrcToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['submitPasswordChangeRequest', [Object.assign({ sessionToken: pcrcToken }, data)]]], null);

    } else if (action === 'passwordChangeRequestUpdate') {
      var pcruToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['updatePasswordChangeRequest', [Object.assign({ sessionToken: pcruToken }, data)]]], null);

    } else if (action === 'rolePermissionsLoad') {
      var rplToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['getRoleButtonPermissions', [rplToken]]], { success: true, data: {} });

    } else if (action === 'rolePermissionsSave') {
      var rpsToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['saveRoleActionPermissions', [Object.assign({ sessionToken: rpsToken }, data)]]], null);

    } else if (action === 'savePortalUser') {
      var spuToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['savePortalUser', [Object.assign({ sessionToken: spuToken }, data)]]], null);

    } else if (action === 'deletePortalUser') {
      var dpuToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['deletePortalUser', [Object.assign({ sessionToken: dpuToken }, data)]]], null);

    } else if (action === 'saveRestaurantRecord') {
      var srrToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['saveRestaurantRecord', [Object.assign({ sessionToken: srrToken }, data)]]], null);

    } else if (action === 'deleteRestaurantRecord') {
      var drrToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['deleteRestaurantRecord', [Object.assign({ sessionToken: drrToken }, data)]]], null);

    } else if (action === 'hostLinksLoad') {
      var hllToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['listHostLinkRecords', [hllToken]]], { success: true, data: { links: [] } });

    } else if (action === 'hostLinkSave') {
      var hlsToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['saveHostLinkRecord', [Object.assign({ sessionToken: hlsToken }, data)]]], null);

    } else if (action === 'hostLinkStatusUpdate') {
      var hlsuToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['updateHostLinkStatus', [Object.assign({ sessionToken: hlsuToken }, data)]]], null);

    } else if (action === 'newslettersLoad') {
      var nlToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['listNewsletterRecords', [nlToken]]], { success: true, data: { newsletters: [] } });

    } else if (action === 'newsletterSave') {
      var nsToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['saveNewsletterRecord', [Object.assign({ sessionToken: nsToken }, data)]]], null);

    } else if (action === 'newsletterStatusUpdate') {
      var nsuToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['updateNewsletterStatus', [Object.assign({ sessionToken: nsuToken }, data)]]], null);

    } else if (action === 'dropdownOptionsLoad') {
      var dolToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['listDropdownOptionRecords', [dolToken]]], { success: true, data: { options: [] } });

    } else if (action === 'dropdownOptionSave') {
      var dosToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['saveDropdownOptionRecord', [Object.assign({ sessionToken: dosToken }, data)]]], null);

    } else if (action === 'dropdownOptionStatusUpdate') {
      var dosuToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['updateDropdownOptionStatus', [Object.assign({ sessionToken: dosuToken }, data)]]], null);

    } else if (action === 'operationsDashboard') {
      var odToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['getOperationsDashboardData', [odToken]]], { success: true, data: {} });

    } else if (action === 'ping') {
      result = { success: true, data: { ok: true } };

    } else if (action === 'initSheets') {
      var isToken = firebaseResolveToken_(data.token);
      result = safeCallOne_([['initializePortalSheetsForCurrentUser', [isToken]]], { success: true, data: {} });

    } else {
      result = { success: false, error: 'Unknown API action: ' + action };
    }

    return firebaseJson_(normalizeApiResult_(result));

  } catch (error) {
    return firebaseJson_({ success: false, error: error.message || String(error) });
  }
}

function firebaseRememberSession_(result) {
  try {
    if (!(result && result.success && result.data && result.data.session)) return result;
    var session = result.data.session;
    var token = session.token || '';
    if (!token) return result;
    var cache = CacheService.getScriptCache();
    cache.put('FIREBASE_SESSION_' + token, JSON.stringify(session), 21600);
    PropertiesService.getScriptProperties().setProperty('FIREBASE_SESSION_' + token, JSON.stringify(session));
    return result;
  } catch (error) { return result; }
}

function firebaseForgetSession_(token) {
  try {
    token = String(token || '').trim();
    if (!token) return;
    CacheService.getScriptCache().remove('FIREBASE_SESSION_' + token);
    PropertiesService.getScriptProperties().deleteProperty('FIREBASE_SESSION_' + token);
  } catch (error) {}
}

function firebaseResolveToken_(token) {
  token = String(token || '').trim();
  if (!token) return token;
  try {
    var cached = CacheService.getScriptCache().get('FIREBASE_SESSION_' + token);
    if (cached) { var s = JSON.parse(cached); return s.token || token; }
  } catch (e1) {}
  try {
    var stored = PropertiesService.getScriptProperties().getProperty('FIREBASE_SESSION_' + token);
    if (stored) { var s2 = JSON.parse(stored); return s2.token || token; }
  } catch (e2) {}
  return token;
}

function firebaseJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function normalizeApiResult_(result) {
  if (result && result.success !== undefined) return result;
  return { success: true, data: result || {} };
}

function safeCallOne_(calls, fallback) {
  var errors = [];
  for (var i = 0; i < calls.length; i++) {
    var name = calls[i][0];
    var args = calls[i][1] || [];
    try {
      var fn = this[name];
      if (typeof fn === 'function') return fn.apply(this, args);
      errors.push(name + ' not found');
    } catch (err) {
      errors.push(name + ': ' + (err.message || String(err)));
    }
  }
  if (fallback) { fallback.debug = { tried: calls.map(function(c){return c[0];}), errors: errors }; return fallback; }
  return { success: false, error: errors.join(' | ') };
}

function firebaseDashboardBoot_(token) {
  var boot = safeCallOne_([
    ['getDashboardBoot', [token]],
    ['loadDashboard', [token]],
    ['getDashboardData', [token]]
  ], null);

  if (boot && boot.success) {
    var data = boot.data || {};
    if (!data.roleActions || !data.roleActions.length) data.roleActions = firebaseDefaultRoleActions_();
    if (!data.hostLinks) data.hostLinks = [];
    if (!data.newsletters) data.newsletters = [];
    boot.data = data;
    return boot;
  }

  return {
    success: true,
    data: {
      hostLinks: [],
      roleActions: firebaseDefaultRoleActions_(),
      newsletters: [],
      metrics: {}
    }
  };
}

function firebaseDefaultRoleActions_() {
  return [
    { key: 'PAYROLL_SUBMIT', label: 'Payroll Submit' },
    { key: 'PAYROLL_REVIEW', label: 'Payroll Review' },
    { key: 'PAYROLL_ADMIN', label: 'Payroll Admin' },
    { key: 'MANAGE_TEAM', label: 'Manage Team' },
    { key: 'RECRUITMENT', label: 'Recruitment' },
    { key: 'RESTAURANT_REQUESTS', label: 'Restaurant Requests' }
  ];
}

function firebaseNormalizeRecruitment_(result) {
  if (!(result && result.success && result.data)) return result;
  var data = result.data;
  data.restaurants = data.restaurants || [];
  data.teamMembers = data.teamMembers || data.members || data.employees || data.newHeroes || [];
  data.members = data.teamMembers;
  data.employees = data.teamMembers;
  data.hiringRequests = data.hiringRequests || data.hiring || data.requests || [];
  return result;
}

function firebaseNormalizePayroll_(result) {
  if (!(result && result.success && result.data)) return result;
  var data = result.data;
  data.restaurants = data.restaurants || [];
  data.employees = data.employees || data.teamMembers || data.members || [];
  data.submissions = data.submissions || data.payrolls || data.rows || [];
  data.approvedPayrolls = data.approvedPayrolls || data.approved || [];
  data.batches = data.batches || data.payrollBatches || [];
  data.roeApplications = data.roeApplications || data.roePool || data.roes || [];
  return result;
}
