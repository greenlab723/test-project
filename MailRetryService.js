function enqueueMailRetry(entry) {
  const sheet = getMailRetrySheet();
  sheet.appendRow([
    new Date(),
    entry.type || '',
    entry.recipient || '',
    entry.subject || '',
    entry.body || '',
    entry.sender || '',
    'pending',
    entry.error || '',
    '',
  ]);
}

function getMailRetryForAdmin(options) {
  const sheet = getMailRetrySheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const createdIndex = header.indexOf('created_at');
  const typeIndex = header.indexOf('type');
  const recipientIndex = header.indexOf('recipient');
  const subjectIndex = header.indexOf('subject');
  const statusIndex = header.indexOf('status');
  const errorIndex = header.indexOf('last_error');

  if ([createdIndex, recipientIndex, subjectIndex, statusIndex].includes(-1)) {
    throw new Error('MailRetry sheet must include created_at, recipient, subject, status columns');
  }

  const query = (options.query || '').toLowerCase();
  const limit = options.limit || 50;
  const page = options.page || 1;
  const offset = (page - 1) * limit;
  const retries = [];
  let matchCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const recipient = row[recipientIndex] || '';
    const subject = row[subjectIndex] || '';
    const type = row[typeIndex] || '';
    const status = row[statusIndex] || '';

    if (query) {
      const haystack = `${recipient} ${subject} ${type} ${status}`.toLowerCase();
      if (!haystack.includes(query)) {
        continue;
      }
    }

    if (matchCount >= offset && retries.length < limit) {
      retries.push({
        rowIndex: i + 1,
        createdAt: row[createdIndex],
        type,
        recipient,
        subject,
        status,
        lastError: row[errorIndex] || '',
      });
    }

    matchCount += 1;
  }

  return { retries, total: matchCount };
}

function retryMailSend(rowIndex) {
  const sheet = getMailRetrySheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const subjectIndex = header.indexOf('subject');
  const bodyIndex = header.indexOf('body');
  const recipientIndex = header.indexOf('recipient');
  const senderIndex = header.indexOf('sender');
  const statusIndex = header.indexOf('status');
  const errorIndex = header.indexOf('last_error');
  const sentIndex = header.indexOf('sent_at');

  if ([subjectIndex, bodyIndex, recipientIndex, statusIndex].includes(-1)) {
    throw new Error('MailRetry sheet is missing required columns');
  }

  const row = data[rowIndex - 1];
  const recipient = row[recipientIndex];
  const subject = row[subjectIndex];
  const body = row[bodyIndex];
  const sender = row[senderIndex];

  try {
    GmailApp.sendEmail(recipient, subject, body, buildSenderOptions(sender));
    sheet.getRange(rowIndex, statusIndex + 1).setValue('sent');
    if (sentIndex !== -1) {
      sheet.getRange(rowIndex, sentIndex + 1).setValue(new Date());
    }
    if (errorIndex !== -1) {
      sheet.getRange(rowIndex, errorIndex + 1).setValue('');
    }
    return true;
  } catch (error) {
    if (errorIndex !== -1) {
      sheet.getRange(rowIndex, errorIndex + 1).setValue(error && error.message ? error.message : String(error));
    }
    sheet.getRange(rowIndex, statusIndex + 1).setValue('failed');
    return false;
  }
}

function processMailRetryQueue() {
  const settings = getSettingsMap();
  const batchSize = Number(settings.ADMIN_RETRY_BATCH_SIZE) || 20;
  const sheet = getMailRetrySheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const statusIndex = header.indexOf('status');

  if (statusIndex === -1) {
    throw new Error('MailRetry sheet must include status column');
  }

  let processed = 0;
  for (let i = 1; i < data.length; i++) {
    if (processed >= batchSize) {
      break;
    }
    if (String(data[i][statusIndex]) !== 'pending') {
      continue;
    }
    retryMailSend(i + 1);
    processed += 1;
  }

  return processed;
}
