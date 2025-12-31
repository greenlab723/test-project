function sendAdminMail(options) {
  const mode = options.mode || 'bulk';
  let subject = options.subject || '';
  let body = options.body || '';
  const settings = getSettingsMap();
  const sender = settings.MAIL_SENDER || '';
  const templateKey = options.templateKey || settings.ADMIN_MAIL_TEMPLATE_KEY || '';

  const template = templateKey ? getMailTemplate(templateKey) : null;
  if (template) {
    if (!subject) {
      subject = template.subject || '';
    }
    if (!body) {
      body = template.body || '';
    }
  }

  if (!subject || !body) {
    throw new Error('件名と本文は必須です。');
  }

  if (mode === 'individual') {
    const targetEmail = options.targetEmail || '';
    if (!targetEmail) {
      throw new Error('送信先メールアドレスを入力してください。');
    }
    const content = template
      ? applyTemplate(template.body, buildAdminTemplateValues(body, '', targetEmail))
      : body;
    sendEmailWithRetry({
      type: 'admin_individual',
      recipient: targetEmail,
      subject,
      body: content,
      sender,
      templateKey,
    });
    return { sentCount: 1 };
  }

  const statusFilter = String(settings.ADMIN_MAIL_STATUS || '').trim();
  const members = getMembersForAdmin({
    query: '',
    status: statusFilter,
    limit: Number(settings.ADMIN_BULK_LIMIT) || 1000,
  }).members;

  members.forEach((member) => {
    if (member.email) {
      const content = template
        ? applyTemplate(template.body, buildAdminTemplateValues(body, member.name, member.email))
        : body;
      sendEmailWithRetry({
        type: 'admin_bulk',
        recipient: member.email,
        subject,
        body: content,
        sender,
        memberId: member.memberId,
        templateKey,
      });
    }
  });

  return { sentCount: members.length };
}

function buildAdminTemplateValues(body, name, email) {
  const settings = getSettingsMap();
  return {
    name: name || '',
    email: email || '',
    body: body || '',
    organization_name: settings.ORGANIZATION_NAME || '',
  };
}
