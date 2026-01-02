function getSettingsMap() {
  const sheet = getSheetByName(SHEET_NAMES.settings);
  const values = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < values.length; i++) {
    const key = values[i][0];
    const value = values[i][1];
    if (key) {
      settings[String(key).trim()] = value;
    }
  }

  return settings;
}

function getSettingValue(key, defaultValue) {
  const settings = getSettingsMap();
  if (Object.prototype.hasOwnProperty.call(settings, key)) {
    return settings[key];
  }
  return defaultValue;
}

function getMembersSheet() {
  return getSheetByName(SHEET_NAMES.members);
}

function getTokensSheet() {
  return getSheetByName(SHEET_NAMES.tokens);
}

function getMailTemplate(templateKey) {
  const sheet = getSheetByName(SHEET_NAMES.mailTemplates);
  const values = sheet.getDataRange().getValues();
  const header = values[0] || [];
  const keyIndex = header.indexOf('template_key');
  const subjectIndex = header.indexOf('subject');
  const bodyIndex = header.indexOf('body');

  if (keyIndex === -1 || subjectIndex === -1 || bodyIndex === -1) {
    throw new Error('MailTemplates sheet must include template_key, subject, body columns');
  }

  for (let i = 1; i < values.length; i++) {
    if (values[i][keyIndex] === templateKey) {
      return {
        subject: values[i][subjectIndex],
        body: values[i][bodyIndex],
      };
    }
  }

  throw new Error(`Mail template not found: ${templateKey}`);
}

function getMailTemplates() {
  const sheet = getSheetByName(SHEET_NAMES.mailTemplates);
  const values = sheet.getDataRange().getValues();
  const header = values[0] || [];
  const keyIndex = header.indexOf('template_key');
  const subjectIndex = header.indexOf('subject');
  const bodyIndex = header.indexOf('body');

  if (keyIndex === -1 || subjectIndex === -1 || bodyIndex === -1) {
    throw new Error('MailTemplates sheet must include template_key, subject, body columns');
  }

  const templates = [];
  for (let i = 1; i < values.length; i++) {
    const key = values[i][keyIndex];
    if (!key) {
      continue;
    }
    templates.push({
      key: String(key),
      subject: values[i][subjectIndex],
      body: values[i][bodyIndex],
    });
  }

  return templates;
}

function getMailTemplateVariables() {
  return [
    { key: 'name', description: '会員名' },
    { key: 'email', description: '会員メールアドレス' },
    { key: 'member_id', description: '会員ID' },
    { key: 'organization_name', description: '団体名' },
    { key: 'verify_url', description: '認証URL（認証メールのみ）' },
    { key: 'body', description: '本文（管理画のメール入力）' },
  ];
}

function getMailLogSheet() {
  return getSheetByName(SHEET_NAMES.mailLog);
}

function getMemberChangeLogSheet() {
  return getSheetByName(SHEET_NAMES.memberChangeLog);
}

function getAdminActionLogSheet() {
  return getSheetByName(SHEET_NAMES.adminActionLog);
}

function getErrorLogSheet() {
  return getSheetByName(SHEET_NAMES.errorLog);
}

function getMailRetrySheet() {
  return getSheetByName(SHEET_NAMES.mailRetry);
}

function getTemplateChangeLogSheet() {
  return getSheetByName(SHEET_NAMES.templateChangeLog);
}

function getFormFields() {
  const sheet = getSheetByName(SHEET_NAMES.formFields);
  const values = sheet.getDataRange().getValues();
  const header = values[0] || [];
  const fieldKeyIndex = header.indexOf('field_key');
  const labelIndex = header.indexOf('label');
  const typeIndex = header.indexOf('type');
  const requiredIndex = header.indexOf('required');
  const orderIndex = header.indexOf('order');

  if ([fieldKeyIndex, labelIndex, typeIndex, requiredIndex, orderIndex].includes(-1)) {
    throw new Error('FormFields sheet must include field_key, label, type, required, order columns');
  }

  const fields = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[fieldKeyIndex]) {
      continue;
    }
    fields.push({
      key: String(row[fieldKeyIndex]).trim(),
      label: row[labelIndex],
      type: row[typeIndex] || 'text',
      required: String(row[requiredIndex]).toLowerCase() === 'true',
      order: Number(row[orderIndex]) || 0,
    });
  }

  return fields.sort((a, b) => a.order - b.order);
}

function getAdminEmails() {
  const sheet = getSheetByName(SHEET_NAMES.adminAccess);
  const values = sheet.getDataRange().getValues();
  const emails = new Set();

  for (let i = 1; i < values.length; i++) {
    const email = values[i][0];
    if (email) {
      emails.add(String(email).trim().toLowerCase());
    }
  }

  return Array.from(emails);
}

function getSheetByName(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error(`${name} sheet not found`);
  }
  return sheet;
}
