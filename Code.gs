function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'register';

  return runWithErrorHandling(() => {
    if (action === 'admin') {
      return renderAdminPage(e);
    }

    if (action === 'admin_dashboard') {
      return renderAdminDashboard(e);
    }

    if (action === 'admin_members_csv') {
      return exportAdminMembersCsv(e);
    }

    if (action === 'admin_mail_log') {
      return renderAdminMailLog(e);
    }

    if (action === 'admin_mail_log_csv') {
      return exportAdminMailLogCsv(e);
    }

    if (action === 'admin_action_log') {
      return renderAdminActionLog(e);
    }

    if (action === 'admin_action_log_csv') {
      return exportAdminActionLogCsv(e);
    }

    if (action === 'admin_change_log') {
      return renderAdminChangeLog(e);
    }

    if (action === 'admin_error_log') {
      return renderAdminErrorLog(e);
    }

    if (action === 'admin_error_log_csv') {
      return exportAdminErrorLogCsv(e);
    }

    if (action === 'admin_template_log') {
      return renderAdminTemplateLog(e);
    }

    if (action === 'admin_mail_retry') {
      return renderAdminMailRetry(e);
    }

    if (action === 'admin_change_log_csv') {
      return exportAdminChangeLogCsv(e);
    }

    if (action === 'admin_templates') {
      return renderAdminTemplates(e);
    }

    if (action === 'admin_member') {
      return renderAdminMemberDetail(e);
    }

    if (action === 'verify') {
      return handleVerification(e);
    }

    if (action === 'update') {
      return renderUpdateForm();
    }

    if (action === 'update_verify') {
      return handleUpdateVerification(e);
    }

    return renderRegisterForm();
  }, { action, params: (e && e.parameter) || {} }, 'エラーが発生しました。');
}

function doPost(e) {
  const action = (e && e.parameter && e.parameter.action) || 'register';

  return runWithErrorHandling(() => {
    if (action === 'admin_send') {
      return handleAdminMailSend(e);
    }

    if (action === 'admin_status') {
      return handleAdminStatusUpdate(e);
    }

    if (action === 'admin_template_save') {
      return handleAdminTemplateSave(e);
    }

    if (action === 'admin_template_restore') {
      return handleAdminTemplateRestore(e);
    }

    if (action === 'admin_retry_send') {
      return handleAdminRetrySend(e);
    }

    if (action === 'update_request') {
      return handleUpdateRequest(e);
    }

    return handleRegistration(e);
  }, { action, params: (e && e.parameter) || {} }, 'エラーが発生しました。');
}

function renderRegisterForm() {
  const settings = getSettingsMap();
  const registrationEnabled = String(settings.REGISTRATION_ENABLED).toLowerCase() === 'true';

  const template = HtmlService.createTemplateFromFile('register');
  template.registrationEnabled = registrationEnabled;
  template.fields = getFormFields();
  template.organizationName = settings.ORGANIZATION_NAME || '';
  return template.evaluate().setTitle('会員登録');
}

function handleRegistration(e) {
  const settings = getSettingsMap();
  const registrationEnabled = String(settings.REGISTRATION_ENABLED).toLowerCase() === 'true';
  if (!registrationEnabled) {
    return renderMessage('登録受付は停止中です。', '会員登録');
  }

  const fields = getFormFields();
  const formData = buildFormData(e, fields);
  const member = registerMember(formData);
  const tokenInfo = createToken(member.memberId, 'register');

  sendRegistrationEmail(member, tokenInfo);

  return renderMessage('仮登録を受け付けました。認証メールをご確認ください。', '会員登録');
}

function handleVerification(e) {
  const token = e && e.parameter && e.parameter.token;
  if (!token) {
    return renderMessage('認証情報が見つかりません。', '会員登録');
  }

  const result = verifyToken(token, 'register');
  if (!result.valid) {
    return renderMessage(result.message, '会員登録');
  }

  activateMember(result.memberId);
  markTokenUsed(token);
  return renderMessage('本登録が完了しました。', '会員登録');
}

function renderMessage(message, title) {
  const template = HtmlService.createTemplateFromFile('register_complete');
  template.message = message;
  return template.evaluate().setTitle(title || '会員登録');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function runWithErrorHandling(fn, context, fallbackMessage) {
  try {
    return fn();
  } catch (error) {
    logError({
      action: context.action,
      userEmail: Session.getActiveUser().getEmail(),
      message: error && error.message ? error.message : String(error),
      stack: error && error.stack ? error.stack : '',
      parameters: JSON.stringify(context.params || {}),
    });
    return renderMessage(fallbackMessage || 'エラーが発生しました。');
  }
}

function renderAdminPage(e) {
  const settings = getSettingsMap();
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみアクセスできます。');
  }

  const query = (e && e.parameter && e.parameter.q) || '';
  const status = (e && e.parameter && e.parameter.status) || '';
  const createdFrom = (e && e.parameter && e.parameter.created_from) || '';
  const createdTo = (e && e.parameter && e.parameter.created_to) || '';
  const updatedFrom = (e && e.parameter && e.parameter.updated_from) || '';
  const updatedTo = (e && e.parameter && e.parameter.updated_to) || '';
  const pageSize =
    Number((e && e.parameter && e.parameter.page_size) || settings.ADMIN_PAGE_SIZE) || 50;
  const page = Math.max(Number((e && e.parameter && e.parameter.page) || 1), 1);

  const result = getMembersForAdmin({
    query,
    status,
    createdFrom,
    createdTo,
    updatedFrom,
    updatedTo,
    limit: pageSize,
    page,
  });
  const templates = getMailTemplates();
  const selectedTemplateKey =
    (e && e.parameter && e.parameter.template_key) || settings.ADMIN_MAIL_TEMPLATE_KEY || '';
  const selectedTemplate = templates.find((template) => template.key === selectedTemplateKey) || null;

  const template = HtmlService.createTemplateFromFile('admin');
  template.members = result.members;
  template.query = query;
  template.status = status;
  template.createdFrom = createdFrom;
  template.createdTo = createdTo;
  template.updatedFrom = updatedFrom;
  template.updatedTo = updatedTo;
  template.pageSize = pageSize;
  template.page = page;
  template.total = result.total;
  template.mailTemplates = templates;
  template.selectedTemplateKey = selectedTemplateKey;
  template.selectedTemplate = selectedTemplate;
  template.templateVariables = getMailTemplateVariables();
  template.organizationName = settings.ORGANIZATION_NAME || '';
  return template.evaluate().setTitle('管理画面');
}

function renderAdminDashboard(e) {
  const settings = getSettingsMap();
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみアクセスできます。', '管理画面');
  }

  const stats = getAdminDashboardStats();
  const template = HtmlService.createTemplateFromFile('admin_dashboard');
  template.stats = stats;
  template.organizationName = settings.ORGANIZATION_NAME || '';
  return template.evaluate().setTitle('ダッシュボード');
}

function exportAdminMembersCsv(e) {
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみアクセスできます。', '管理画面');
  }

  const query = (e && e.parameter && e.parameter.q) || '';
  const status = (e && e.parameter && e.parameter.status) || '';
  const createdFrom = (e && e.parameter && e.parameter.created_from) || '';
  const createdTo = (e && e.parameter && e.parameter.created_to) || '';
  const updatedFrom = (e && e.parameter && e.parameter.updated_from) || '';
  const updatedTo = (e && e.parameter && e.parameter.updated_to) || '';
  const fieldKeys = getFormFields().map((field) => field.key);

  const rows = getMemberRowsForExport({
    query,
    status,
    createdFrom,
    createdTo,
    updatedFrom,
    updatedTo,
    fieldKeys,
  });
  const header = ['member_id', 'name', 'email', 'status', 'created_at', 'updated_at', ...fieldKeys];
  return createCsvResponse(header, rows, 'members.csv');
}

function renderAdminMailLog(e) {
  const settings = getSettingsMap();
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみアクセスできます。', '管理画面');
  }

  const query = (e && e.parameter && e.parameter.q) || '';
  const limit = Number((e && e.parameter && e.parameter.page_size) || settings.ADMIN_MAILLOG_PAGE_SIZE) || 50;
  const page = Math.max(Number((e && e.parameter && e.parameter.page) || 1), 1);

  const result = getMailLogForAdmin({ query, limit, page });

  const template = HtmlService.createTemplateFromFile('admin_mail_log');
  template.logs = result.logs;
  template.query = query;
  template.pageSize = limit;
  template.page = page;
  template.total = result.total;
  template.organizationName = settings.ORGANIZATION_NAME || '';
  return template.evaluate().setTitle('メール履歴');
}

function exportAdminMailLogCsv(e) {
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみアクセスできます。', '管理画面');
  }

  const query = (e && e.parameter && e.parameter.q) || '';
  const rows = getMailLogRowsForExport({ query });
  const header = ['sent_at', 'type', 'subject', 'recipient', 'template_key', 'sender'];
  return createCsvResponse(header, rows, 'mail_log.csv');
}

function renderAdminActionLog(e) {
  const settings = getSettingsMap();
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみアクセスできます。', '管理画面');
  }

  const query = (e && e.parameter && e.parameter.q) || '';
  const limit =
    Number((e && e.parameter && e.parameter.page_size) || settings.ADMIN_ACTIONLOG_PAGE_SIZE) || 50;
  const page = Math.max(Number((e && e.parameter && e.parameter.page) || 1), 1);

  const result = getAdminActionLogForAdmin({ query, limit, page });

  const template = HtmlService.createTemplateFromFile('admin_action_log');
  template.logs = result.logs;
  template.query = query;
  template.pageSize = limit;
  template.page = page;
  template.total = result.total;
  template.organizationName = settings.ORGANIZATION_NAME || '';
  return template.evaluate().setTitle('操作ログ');
}

function exportAdminActionLogCsv(e) {
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみアクセスできます。', '管理画面');
  }

  const query = (e && e.parameter && e.parameter.q) || '';
  const rows = getAdminActionLogRowsForExport({ query });
  const header = ['action_at', 'admin_email', 'action', 'target', 'details'];
  return createCsvResponse(header, rows, 'admin_action_log.csv');
}

function renderAdminChangeLog(e) {
  const settings = getSettingsMap();
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみアクセスできます。', '管理画面');
  }

  const query = (e && e.parameter && e.parameter.q) || '';
  const memberId = (e && e.parameter && e.parameter.member_id) || '';
  const limit =
    Number((e && e.parameter && e.parameter.page_size) || settings.ADMIN_CHANGELOG_PAGE_SIZE) || 50;
  const page = Math.max(Number((e && e.parameter && e.parameter.page) || 1), 1);

  const result = getMemberChangeLogForAdmin({ query, memberId, limit, page });

  const template = HtmlService.createTemplateFromFile('admin_change_log');
  template.logs = result.logs;
  template.query = query;
  template.memberId = memberId;
  template.pageSize = limit;
  template.page = page;
  template.total = result.total;
  template.organizationName = settings.ORGANIZATION_NAME || '';
  return template.evaluate().setTitle('変更履歴');
}

function renderAdminErrorLog(e) {
  const settings = getSettingsMap();
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみアクセスできます。', '管理画面');
  }

  const query = (e && e.parameter && e.parameter.q) || '';
  const limit =
    Number((e && e.parameter && e.parameter.page_size) || settings.ADMIN_ERRORLOG_PAGE_SIZE) || 50;
  const page = Math.max(Number((e && e.parameter && e.parameter.page) || 1), 1);

  const result = getErrorLogForAdmin({ query, limit, page });

  const template = HtmlService.createTemplateFromFile('admin_error_log');
  template.logs = result.logs;
  template.query = query;
  template.pageSize = limit;
  template.page = page;
  template.total = result.total;
  template.organizationName = settings.ORGANIZATION_NAME || '';
  return template.evaluate().setTitle('エラーログ');
}

function exportAdminErrorLogCsv(e) {
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみアクセスできます。', '管理画面');
  }

  const query = (e && e.parameter && e.parameter.q) || '';
  const rows = getErrorLogRowsForExport({ query });
  const header = ['error_at', 'action', 'user_email', 'message', 'stack', 'parameters'];
  return createCsvResponse(header, rows, 'error_log.csv');
}

function renderAdminTemplateLog(e) {
  const settings = getSettingsMap();
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみアクセスできます。', '管理画面');
  }

  const query = (e && e.parameter && e.parameter.q) || '';
  const limit =
    Number((e && e.parameter && e.parameter.page_size) || settings.ADMIN_TEMPLATELOG_PAGE_SIZE) || 50;
  const page = Math.max(Number((e && e.parameter && e.parameter.page) || 1), 1);

  const result = getTemplateChangeLogForAdmin({ query, limit, page });

  const template = HtmlService.createTemplateFromFile('admin_template_log');
  template.logs = result.logs;
  template.query = query;
  template.pageSize = limit;
  template.page = page;
  template.total = result.total;
  template.organizationName = settings.ORGANIZATION_NAME || '';
  return template.evaluate().setTitle('テンプレート変更履歴');
}

function renderAdminMailRetry(e) {
  const settings = getSettingsMap();
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみアクセスできます。', '管理画面');
  }

  const query = (e && e.parameter && e.parameter.q) || '';
  const limit =
    Number((e && e.parameter && e.parameter.page_size) || settings.ADMIN_RETRY_PAGE_SIZE) || 50;
  const page = Math.max(Number((e && e.parameter && e.parameter.page) || 1), 1);

  const result = getMailRetryForAdmin({ query, limit, page });

  const template = HtmlService.createTemplateFromFile('admin_mail_retry');
  template.retries = result.retries;
  template.query = query;
  template.pageSize = limit;
  template.page = page;
  template.total = result.total;
  template.organizationName = settings.ORGANIZATION_NAME || '';
  return template.evaluate().setTitle('再送キュー');
}

function exportAdminChangeLogCsv(e) {
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみアクセスできます。', '管理画面');
  }

  const query = (e && e.parameter && e.parameter.q) || '';
  const rows = getMemberChangeLogRowsForExport({ query });
  const header = ['changed_at', 'member_id', 'source', 'changes'];
  return createCsvResponse(header, rows, 'member_change_log.csv');
}

function createCsvResponse(header, rows, filename) {
  const csv = [header, ...rows].map((row) => row.map(escapeCsvValue).join(',')).join('\n');
  return ContentService.createTextOutput(csv)
    .setMimeType(ContentService.MimeType.CSV)
    .setFileName(filename);
}

function escapeCsvValue(value) {
  if (value === null || value === undefined) {
    return '';
  }
  const text = String(value);
  if (text.includes('"') || text.includes(',') || text.includes('\n')) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function renderAdminTemplates(e) {
  const settings = getSettingsMap();
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみアクセスできます。', '管理画面');
  }

  const templates = getMailTemplates();
  const template = HtmlService.createTemplateFromFile('admin_templates');
  template.templates = templates;
  template.organizationName = settings.ORGANIZATION_NAME || '';
  return template.evaluate().setTitle('テンプレート管理');
}

function renderAdminMemberDetail(e) {
  const settings = getSettingsMap();
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみアクセスできます。', '管理画面');
  }

  const memberId = (e && e.parameter && e.parameter.member_id) || '';
  if (!memberId) {
    return renderMessage('会員IDが指定されていません。', '管理画面');
  }

  const member = getMemberDetailForAdmin(memberId);
  const template = HtmlService.createTemplateFromFile('admin_member_detail');
  template.member = member;
  template.organizationName = settings.ORGANIZATION_NAME || '';
  return template.evaluate().setTitle('会員詳細');
}

function handleAdminMailSend(e) {
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみ操作できます。', '管理画面');
  }

  const params = (e && e.parameter) || {};
  const mode = params.mode || 'bulk';
  const subject = params.subject || '';
  const body = params.body || '';
  const templateKey = params.template_key || '';
  const targetEmail = params.target_email || '';

  if (mode === 'individual' && !targetEmail) {
    return renderMessage('送信先メールアドレスを入力してください。', '管理画面');
  }

  const result = sendAdminMail({
    mode,
    subject,
    body,
    targetEmail,
    templateKey,
  });

  logAdminAction({
    adminEmail: activeUser,
    action: 'mail_send',
    target: mode === 'individual' ? targetEmail : 'bulk',
    details: `template=${templateKey || 'default'} sent=${result.sentCount}`,
  });

  return renderMessage(`メール送信を完了しました。送信件数: ${result.sentCount}件`, '管理画面');
}

function handleAdminStatusUpdate(e) {
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみ操作できます。', '管理画面');
  }

  const params = (e && e.parameter) || {};
  const memberId = params.member_id || '';
  const status = params.status || '';

  if (!memberId || !status) {
    return renderMessage('会員IDとステータスを指定してください。', '管理画面');
  }

  updateMemberStatus(memberId, status);
  logAdminAction({
    adminEmail: activeUser,
    action: 'status_update',
    target: memberId,
    details: `status=${status}`,
  });
  return renderMessage('ステータスを更新しました。', '管理画面');
}

function handleAdminTemplateSave(e) {
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみ操作できます。', '管理画面');
  }

  const params = (e && e.parameter) || {};
  const templateKey = params.template_key || '';
  const subject = params.subject || '';
  const body = params.body || '';

  if (!templateKey) {
    return renderMessage('テンプレートキーを入力してください。', 'テンプレート管理');
  }

  saveMailTemplate({ templateKey, subject, body, adminEmail: activeUser });
  logAdminAction({
    adminEmail: activeUser,
    action: 'template_save',
    target: templateKey,
    details: 'mail template updated',
  });

  return renderMessage('テンプレートを保存しました。', 'テンプレート管理');
}

function handleAdminTemplateRestore(e) {
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみ操作できます。', 'テンプレート管理');
  }

  const params = (e && e.parameter) || {};
  const templateKey = params.template_key || '';
  const subject = params.subject || '';
  const body = params.body || '';

  if (!templateKey) {
    return renderMessage('テンプレートキーが不足しています。', 'テンプレート管理');
  }

  saveMailTemplate({ templateKey, subject, body, adminEmail: activeUser });
  logAdminAction({
    adminEmail: activeUser,
    action: 'template_restore',
    target: templateKey,
    details: 'template restored',
  });

  return renderMessage('テンプレートを復元しました。', 'テンプレート管理');
}

function handleAdminRetrySend(e) {
  const adminEmails = getAdminEmails();
  const activeUser = Session.getActiveUser().getEmail();
  const isAdmin = activeUser && adminEmails.includes(activeUser.toLowerCase());

  if (!isAdmin) {
    return renderMessage('管理者のみ操作できます。', '管理画面');
  }

  const params = (e && e.parameter) || {};
  const rowIndex = Number(params.row_index || 0);
  if (!rowIndex) {
    return renderMessage('再送対象が見つかりません。', '再送キュー');
  }

  const success = retryMailSend(rowIndex);
  if (success) {
    logAdminAction({
      adminEmail: activeUser,
      action: 'mail_retry',
      target: String(rowIndex),
      details: 'retry success',
    });
    return renderMessage('再送が完了しました。', '再送キュー');
  }

  logAdminAction({
    adminEmail: activeUser,
    action: 'mail_retry',
    target: String(rowIndex),
    details: 'retry failed',
  });
  return renderMessage('再送に失敗しました。', '再送キュー');
}

function renderUpdateForm() {
  const settings = getSettingsMap();
  const updateEnabled = String(settings.UPDATE_ENABLED).toLowerCase() === 'true';

  const template = HtmlService.createTemplateFromFile('update');
  template.updateEnabled = updateEnabled;
  template.fields = getFormFields();
  template.organizationName = settings.ORGANIZATION_NAME || '';
  return template.evaluate().setTitle('会員情報変更');
}

function handleUpdateRequest(e) {
  const settings = getSettingsMap();
  const updateEnabled = String(settings.UPDATE_ENABLED).toLowerCase() === 'true';
  if (!updateEnabled) {
    return renderMessage('会員情報の変更は停止中です。', '会員情報変更');
  }

  const currentEmail = (e && e.parameter && e.parameter.current_email) || '';
  if (!currentEmail) {
    return renderMessage('現在のメールアドレスを入力してください。', '会員情報変更');
  }

  const member = findMemberByEmail(currentEmail);
  if (!member) {
    return renderMessage('該当する会員が見つかりませんでした。', '会員情報変更');
  }

  const fields = getFormFields();
  const formData = buildFormData(e, fields);
  const tokenInfo = createToken(member.memberId, 'update', formData);

  sendUpdateEmail(member, tokenInfo);

  return renderMessage('変更確認メールを送信しました。メールをご確認ください。', '会員情報変更');
}

function handleUpdateVerification(e) {
  const token = e && e.parameter && e.parameter.token;
  if (!token) {
    return renderMessage('認証情報が見つかりません。', '会員情報変更');
  }

  const result = verifyToken(token, 'update');
  if (!result.valid) {
    return renderMessage(result.message, '会員情報変更');
  }

  if (!result.payload) {
    return renderMessage('変更内容が見つかりませんでした。', '会員情報変更');
  }

  updateMember(result.memberId, result.payload);
  markTokenUsed(token);
  return renderMessage('会員情報の変更が完了しました。', '会員情報変更');
}