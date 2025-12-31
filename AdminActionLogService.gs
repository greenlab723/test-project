function logAdminAction(entry) {
  const sheet = getAdminActionLogSheet();
  sheet.appendRow([
    new Date(),
    entry.adminEmail || '',
    entry.action || '',
    entry.target || '',
    entry.details || '',
  ]);
}

function getAdminActionLogForAdmin(options) {
  const sheet = getAdminActionLogSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const actionAtIndex = header.indexOf('action_at');
  const adminIndex = header.indexOf('admin_email');
  const actionIndex = header.indexOf('action');
  const targetIndex = header.indexOf('target');
  const detailsIndex = header.indexOf('details');

  if ([actionAtIndex, adminIndex, actionIndex].includes(-1)) {
    throw new Error('AdminActionLog sheet must include action_at, admin_email, action columns');
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
    const action = row[actionIndex] || '';
    const target = row[targetIndex] || '';
    const details = row[detailsIndex] || '';

    if (query) {
      const haystack = `${adminEmail} ${action} ${target} ${details}`.toLowerCase();
      if (!haystack.includes(query)) {
        continue;
      }
    }

    if (matchCount >= offset && logs.length < limit) {
      logs.push({
        actionAt: row[actionAtIndex],
        adminEmail,
        action,
        target,
        details,
      });
    }

    matchCount += 1;
  }

  return { logs, total: matchCount };
}

function getAdminActionLogRowsForExport(options) {
  const sheet = getAdminActionLogSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const actionAtIndex = header.indexOf('action_at');
  const adminIndex = header.indexOf('admin_email');
  const actionIndex = header.indexOf('action');
  const targetIndex = header.indexOf('target');
  const detailsIndex = header.indexOf('details');

  if ([actionAtIndex, adminIndex, actionIndex].includes(-1)) {
    throw new Error('AdminActionLog sheet must include action_at, admin_email, action columns');
  }

  const query = (options.query || '').toLowerCase();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const adminEmail = row[adminIndex] || '';
    const action = row[actionIndex] || '';
    const target = row[targetIndex] || '';
    const details = row[detailsIndex] || '';

    if (query) {
      const haystack = `${adminEmail} ${action} ${target} ${details}`.toLowerCase();
      if (!haystack.includes(query)) {
        continue;
      }
    }

    rows.push([row[actionAtIndex], adminEmail, action, target, details]);
  }

  return rows;
}
