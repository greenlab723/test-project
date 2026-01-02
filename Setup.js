function initializeSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  createSheetIfMissing(spreadsheet, SHEET_NAMES.settings, ['KEY', 'VALUE']);
  createSheetIfMissing(
    spreadsheet,
    SHEET_NAMES.members,
    ['member_id', 'name', 'email', 'status', 'created_at', 'updated_at', 'form_data']
  );
  createSheetIfMissing(
    spreadsheet,
    SHEET_NAMES.tokens,
    ['token', 'member_id', 'type', 'expires_at', 'created_at', 'used_at', 'payload']
  );
  createSheetIfMissing(
    spreadsheet,
    SHEET_NAMES.mailTemplates,
    ['template_key', 'subject', 'body']
  );
  createSheetIfMissing(
    spreadsheet,
    SHEET_NAMES.formFields,
    ['field_key', 'label', 'type', 'required', 'order']
  );
  createSheetIfMissing(spreadsheet, SHEET_NAMES.adminAccess, ['email']);
  createSheetIfMissing(
    spreadsheet,
    SHEET_NAMES.mailLog,
    ['sent_at', 'type', 'subject', 'recipient', 'member_id', 'template_key', 'sender', 'status']
  );
  createSheetIfMissing(
    spreadsheet,
    SHEET_NAMES.memberChangeLog,
    ['changed_at', 'member_id', 'source', 'changes']
  );
  createSheetIfMissing(
    spreadsheet,
    SHEET_NAMES.adminActionLog,
    ['action_at', 'admin_email', 'action', 'target', 'details']
  );
  createSheetIfMissing(
    spreadsheet,
    SHEET_NAMES.errorLog,
    ['error_at', 'action', 'user_email', 'message', 'stack', 'parameters']
  );
  createSheetIfMissing(
    spreadsheet,
    SHEET_NAMES.mailRetry,
    ['created_at', 'type', 'recipient', 'subject', 'body', 'sender', 'status', 'last_error', 'sent_at']
  );
  createSheetIfMissing(
    spreadsheet,
    SHEET_NAMES.templateChangeLog,
    ['changed_at', 'admin_email', 'template_key', 'before_subject', 'after_subject', 'before_body', 'after_body']
  );

  seedSettings(spreadsheet.getSheetByName(SHEET_NAMES.settings));
  seedMailTemplates(spreadsheet.getSheetByName(SHEET_NAMES.mailTemplates));
  seedFormFields(spreadsheet.getSheetByName(SHEET_NAMES.formFields));
  ensureSheetColumns(
    spreadsheet.getSheetByName(SHEET_NAMES.tokens),
    ['token', 'member_id', 'type', 'expires_at', 'created_at', 'used_at', 'payload']
  );
}

function createSheetIfMissing(spreadsheet, name, header) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    sheet.appendRow(header);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(header);
  }
}

function seedSettings(sheet) {
  const defaults = [
    ['REGISTRATION_ENABLED', 'TRUE'],
    ['UPDATE_ENABLED', 'TRUE'],
    ['TOKEN_EXPIRE_MINUTES', '60'],
    ['ORGANIZATION_NAME', ''],
    ['WEB_APP_URL', ''],
    ['MAIL_SENDER', ''],
    ['ADMIN_PAGE_SIZE', '50'],
    ['ADMIN_BULK_LIMIT', '1000'],
    ['ADMIN_MAIL_STATUS', 'active'],
    ['ADMIN_MAIL_TEMPLATE_KEY', 'admin_bulk'],
    ['ADMIN_MAILLOG_PAGE_SIZE', '50'],
    ['ADMIN_ACTIONLOG_PAGE_SIZE', '50'],
    ['ADMIN_CHANGELOG_PAGE_SIZE', '50'],
    ['ADMIN_ERRORLOG_PAGE_SIZE', '50'],
    ['ADMIN_RETRY_PAGE_SIZE', '50'],
    ['ADMIN_RETRY_BATCH_SIZE', '20'],
    ['ADMIN_TEMPLATELOG_PAGE_SIZE', '50'],
  ];
  appendDefaults(sheet, defaults, 0);
}

function seedMailTemplates(sheet) {
  const defaults = [
    [
      'register_verify',
      '【{{organization_name}}】メール認証のお願い',
      'こんにちは、{{name}}様。\n\n以下のリンクをクリックして本登録を完了してください。\n{{verify_url}}\n\n※このリンクは一定時間で無効になります。',
    ],
    [
      'update_verify',
      '【{{organization_name}}】会員情報変更の確認',
      'こんにちは、{{name}}様。\n\n以下のリンクをクリックして会員情報の変更を完了してください。\n{{verify_url}}\n\n※このリンクは一定時間で無効になります。',
    ],
    [
      'admin_bulk',
      '【{{organization_name}}】お知らせ',
      '{{name}}様\n\n{{body}}\n',
    ],
  ];
  appendDefaults(sheet, defaults, 0);
}

function seedFormFields(sheet) {
  const defaults = [
    ['name', 'お名前', 'text', 'TRUE', 1],
    ['email', 'メールアドレス', 'email', 'TRUE', 2],
  ];
  appendDefaults(sheet, defaults, 0);
}

function appendDefaults(sheet, rows, keyIndex) {
  const existing = sheet.getDataRange().getValues();
  const headerOffset = existing.length > 0 ? 1 : 0;
  const existingKeys = new Set();

  for (let i = 1; i < existing.length; i++) {
    const key = existing[i][keyIndex];
    if (key) {
      existingKeys.add(String(key));
    }
  }

  rows.forEach((row) => {
    if (!existingKeys.has(String(row[keyIndex]))) {
      sheet.appendRow(row);
    }
  });
}

function ensureSheetColumns(sheet, columns) {
  if (!sheet) {
    return;
  }
  const header = sheet.getDataRange().getValues()[0] || [];
  const missing = columns.filter((column) => !header.includes(column));
  if (missing.length === 0) {
    return;
  }
  missing.forEach((column) => header.push(column));
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
}
