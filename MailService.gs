function sendRegistrationEmail(member, tokenInfo) {
  const template = getMailTemplate('register_verify');
  const settings = getSettingsMap();
  const webAppUrl = settings.WEB_APP_URL || ScriptApp.getService().getUrl();
  const verifyUrl = `${webAppUrl}?action=verify&token=${encodeURIComponent(tokenInfo.token)}`;
  const sender = settings.MAIL_SENDER || '';

  const replacements = {
    name: member.name,
    email: member.email,
    member_id: member.memberId,
    verify_url: verifyUrl,
    organization_name: settings.ORGANIZATION_NAME || '',
  };

  const subject = applyTemplate(template.subject, replacements);
  const body = applyTemplate(template.body, replacements);

  sendEmailWithRetry({
    type: 'register_verify',
    recipient: member.email,
    subject,
    body,
    sender,
    memberId: member.memberId,
    templateKey: 'register_verify',
  });
}

function sendUpdateEmail(member, tokenInfo) {
  const template = getMailTemplate('update_verify');
  const settings = getSettingsMap();
  const webAppUrl = settings.WEB_APP_URL || ScriptApp.getService().getUrl();
  const verifyUrl = `${webAppUrl}?action=update_verify&token=${encodeURIComponent(tokenInfo.token)}`;
  const sender = settings.MAIL_SENDER || '';

  const replacements = {
    name: member.name,
    email: member.email,
    member_id: member.memberId,
    verify_url: verifyUrl,
    organization_name: settings.ORGANIZATION_NAME || '',
  };

  const subject = applyTemplate(template.subject, replacements);
  const body = applyTemplate(template.body, replacements);

  sendEmailWithRetry({
    type: 'update_verify',
    recipient: member.email,
    subject,
    body,
    sender,
    memberId: member.memberId,
    templateKey: 'update_verify',
  });
}

function applyTemplate(content, replacements) {
  return Object.keys(replacements).reduce((result, key) => {
    const pattern = new RegExp(`{{\\s*${key}\\s*}}`, 'g');
    return result.replace(pattern, replacements[key] || '');
  }, content || '');
}

function logMailSend(entry) {
  const sheet = getMailLogSheet();
  const row = [
    new Date(),
    entry.type || '',
    entry.subject || '',
    entry.recipient || '',
    entry.memberId || '',
    entry.templateKey || '',
    entry.sender || '',
    entry.status || '',
  ];
  sheet.appendRow(row);
}

function buildSenderOptions(sender) {
  if (!sender) {
    return {};
  }
  return { from: sender };
}

function sendEmailWithRetry(entry) {
  try {
    GmailApp.sendEmail(entry.recipient, entry.subject, entry.body, buildSenderOptions(entry.sender));
    logMailSend({
      type: entry.type,
      subject: entry.subject,
      recipient: entry.recipient,
      memberId: entry.memberId || '',
      templateKey: entry.templateKey || '',
      sender: entry.sender || '',
      status: 'sent',
    });
  } catch (error) {
    logMailSend({
      type: entry.type,
      subject: entry.subject,
      recipient: entry.recipient,
      memberId: entry.memberId || '',
      templateKey: entry.templateKey || '',
      sender: entry.sender || '',
      status: 'failed',
    });
    enqueueMailRetry({
      type: entry.type,
      recipient: entry.recipient,
      subject: entry.subject,
      body: entry.body,
      sender: entry.sender,
      error: error && error.message ? error.message : String(error),
    });
  }
}
