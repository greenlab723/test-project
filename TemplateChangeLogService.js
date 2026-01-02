function logTemplateChange(entry) {
  const sheet = getTemplateChangeLogSheet();
  sheet.appendRow([
    new Date(),
    entry.adminEmail || '',
    entry.templateKey || '',
    entry.beforeSubject || '',
    entry.afterSubject || '',
    entry.beforeBody || '',
    entry.afterBody || '',
  ]);
}

function getTemplateChangeLogForAdmin(options) {
  const sheet = getTemplateChangeLogSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const changedIndex = header.indexOf('changed_at');
  const adminIndex = header.indexOf('admin_email');
  const keyIndex = header.indexOf('template_key');
  const beforeSubjectIndex = header.indexOf('before_subject');
  const afterSubjectIndex = header.indexOf('after_subject');
  const beforeBodyIndex = header.indexOf('before_body');
  const afterBodyIndex = header.indexOf('after_body');

  if ([changedIndex, adminIndex, keyIndex].includes(-1)) {
    throw new Error('TemplateChangeLog sheet must include changed_at, admin_email, template_key');
  }

  const query = (options.query || '').toLowerCase();
  const limit = options.limit || 50;
  const page = options.page || 1;
  const offset = (page - 1) * limit;
  const logs = [];
  let matchCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const adminEmail = row[adminIndex] || '';
    const templateKey = row[keyIndex] || '';

    if (query) {
      const haystack = `${adminEmail} ${templateKey}`.toLowerCase();
      if (!haystack.includes(query)) {
        continue;
      }
    }

    if (matchCount >= offset && logs.length < limit) {
      logs.push({
        changedAt: row[changedIndex],
        adminEmail,
        templateKey,
        beforeSubject: row[beforeSubjectIndex] || '',
        afterSubject: row[afterSubjectIndex] || '',
        beforeBody: row[beforeBodyIndex] || '',
        afterBody: row[afterBodyIndex] || '',
      });
    }

    matchCount += 1;
  }

  return { logs, total: matchCount };
}
