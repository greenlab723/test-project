function logError(entry) {
  const sheet = getErrorLogSheet();
  sheet.appendRow([
    new Date(),
    entry.action || '',
    entry.userEmail || '',
    entry.message || '',
    entry.stack || '',
    entry.parameters || '',
  ]);
}

function getErrorLogForAdmin(options) {
  const sheet = getErrorLogSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const errorAtIndex = header.indexOf('error_at');
  const actionIndex = header.indexOf('action');
  const userIndex = header.indexOf('user_email');
  const messageIndex = header.indexOf('message');
  const stackIndex = header.indexOf('stack');
  const paramsIndex = header.indexOf('parameters');

  if ([errorAtIndex, actionIndex, userIndex, messageIndex].includes(-1)) {
    throw new Error('ErrorLog sheet must include error_at, action, user_email, message columns');
  }

  const query = (options.query || '').toLowerCase();
  const limit = options.limit || 50;
  const page = options.page || 1;
  const offset = (page - 1) * limit;
  const logs = [];
  let matchCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const action = row[actionIndex] || '';
    const userEmail = row[userIndex] || '';
    const message = row[messageIndex] || '';
    const stack = row[stackIndex] || '';
    const parameters = row[paramsIndex] || '';

    if (query) {
      const haystack = `${action} ${userEmail} ${message}`.toLowerCase();
      if (!haystack.includes(query)) {
        continue;
      }
    }

    if (matchCount >= offset && logs.length < limit) {
      logs.push({
        errorAt: row[errorAtIndex],
        action,
        userEmail,
        message,
        stack,
        parameters,
      });
    }

    matchCount += 1;
  }

  return { logs, total: matchCount };
}

function getErrorLogRowsForExport(options) {
  const sheet = getErrorLogSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const errorAtIndex = header.indexOf('error_at');
  const actionIndex = header.indexOf('action');
  const userIndex = header.indexOf('user_email');
  const messageIndex = header.indexOf('message');
  const stackIndex = header.indexOf('stack');
  const paramsIndex = header.indexOf('parameters');

  if ([errorAtIndex, actionIndex, userIndex, messageIndex].includes(-1)) {
    throw new Error('ErrorLog sheet must include error_at, action, user_email, message columns');
  }

  const query = (options.query || '').toLowerCase();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const action = row[actionIndex] || '';
    const userEmail = row[userIndex] || '';
    const message = row[messageIndex] || '';

    if (query) {
      const haystack = `${action} ${userEmail} ${message}`.toLowerCase();
      if (!haystack.includes(query)) {
        continue;
      }
    }

    rows.push([
      row[errorAtIndex],
      action,
      userEmail,
      message,
      row[stackIndex] || '',
      row[paramsIndex] || '',
    ]);
  }

  return rows;
}
