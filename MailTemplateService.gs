function saveMailTemplate(template) {
  const sheet = getSheetByName(SHEET_NAMES.mailTemplates);
  const values = sheet.getDataRange().getValues();
  const header = values[0] || [];
  const keyIndex = header.indexOf('template_key');
  const subjectIndex = header.indexOf('subject');
  const bodyIndex = header.indexOf('body');

  if (keyIndex === -1 || subjectIndex === -1 || bodyIndex === -1) {
    throw new Error('MailTemplates sheet must include template_key, subject, body columns');
  }

  const templateKey = template.templateKey;
  let beforeSubject = '';
  let beforeBody = '';
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][keyIndex]) === String(templateKey)) {
      beforeSubject = values[i][subjectIndex] || '';
      beforeBody = values[i][bodyIndex] || '';
      sheet.getRange(i + 1, subjectIndex + 1).setValue(template.subject || '');
      sheet.getRange(i + 1, bodyIndex + 1).setValue(template.body || '');
      logTemplateChange({
        adminEmail: template.adminEmail || '',
        templateKey,
        beforeSubject,
        afterSubject: template.subject || '',
        beforeBody,
        afterBody: template.body || '',
      });
      return;
    }
  }

  const row = new Array(header.length).fill('');
  row[keyIndex] = templateKey;
  row[subjectIndex] = template.subject || '';
  row[bodyIndex] = template.body || '';
  sheet.appendRow(row);
  logTemplateChange({
    adminEmail: template.adminEmail || '',
    templateKey,
    beforeSubject: '',
    afterSubject: template.subject || '',
    beforeBody: '',
    afterBody: template.body || '',
  });
}
