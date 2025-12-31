function getMailLogForAdmin(options) {
  const sheet = getMailLogSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const sentIndex = header.indexOf('sent_at');
  const subjectIndex = header.indexOf('subject');
  const recipientIndex = header.indexOf('recipient');
  const typeIndex = header.indexOf('type');
  const templateIndex = header.indexOf('template_key');
  const senderIndex = header.indexOf('sender');

  if ([sentIndex, subjectIndex, recipientIndex].includes(-1)) {
    throw new Error('MailLog sheet must include sent_at, subject, recipient columns');
  }

  const query = (options.query || '').toLowerCase();
  const limit = options.limit || 50;
  const page = options.page || 1;
  const offset = (page - 1) * limit;
  const logs = [];
  let matchCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const subject = row[subjectIndex] || '';
    const recipient = row[recipientIndex] || '';
    const type = row[typeIndex] || '';

    if (query) {
      const haystack = `${subject} ${recipient} ${type}`.toLowerCase();
      if (!haystack.includes(query)) {
        continue;
      }
    }

    if (matchCount >= offset && logs.length < limit) {
      logs.push({
        sentAt: row[sentIndex],
        subject,
        recipient,
        type,
        templateKey: row[templateIndex] || '',
        sender: row[senderIndex] || '',
      });
    }

    matchCount += 1;
  }

  return { logs, total: matchCount };
}

function getMailLogRowsForExport(options) {
  const sheet = getMailLogSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const sentIndex = header.indexOf('sent_at');
  const subjectIndex = header.indexOf('subject');
  const recipientIndex = header.indexOf('recipient');
  const typeIndex = header.indexOf('type');
  const templateIndex = header.indexOf('template_key');
  const senderIndex = header.indexOf('sender');

  if ([sentIndex, subjectIndex, recipientIndex].includes(-1)) {
    throw new Error('MailLog sheet must include sent_at, subject, recipient columns');
  }

  const query = (options.query || '').toLowerCase();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const subject = row[subjectIndex] || '';
    const recipient = row[recipientIndex] || '';
    const type = row[typeIndex] || '';

    if (query) {
      const haystack = `${subject} ${recipient} ${type}`.toLowerCase();
      if (!haystack.includes(query)) {
        continue;
      }
    }

    rows.push([
      row[sentIndex],
      type,
      subject,
      recipient,
      row[templateIndex] || '',
      row[senderIndex] || '',
    ]);
  }

  return rows;
}
