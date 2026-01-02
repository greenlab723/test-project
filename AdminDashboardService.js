function getAdminDashboardStats() {
  const memberStats = getMemberStatusCounts();
  const mailStats = getRecentMailCounts(7);

  return {
    members: memberStats,
    mail: mailStats,
  };
}

function getMemberStatusCounts() {
  const sheet = getMembersSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const statusIndex = header.indexOf('status');

  if (statusIndex === -1) {
    throw new Error('Members sheet must include status column');
  }

  const counts = {
    active: 0,
    inactive: 0,
    pending: 0,
    total: 0,
  };

  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][statusIndex] || '');
    if (status && Object.prototype.hasOwnProperty.call(counts, status)) {
      counts[status] += 1;
    }
    counts.total += 1;
  }

  return counts;
}

function getRecentMailCounts(days) {
  const sheet = getMailLogSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const sentIndex = header.indexOf('sent_at');
  const typeIndex = header.indexOf('type');

  if (sentIndex === -1 || typeIndex === -1) {
    throw new Error('MailLog sheet must include sent_at and type columns');
  }

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);

  const counts = {
    total: 0,
    register_verify: 0,
    update_verify: 0,
    admin_bulk: 0,
    admin_individual: 0,
  };

  for (let i = 1; i < data.length; i++) {
    const sentAt = data[i][sentIndex];
    if (!sentAt) {
      continue;
    }
    const sentDate = new Date(sentAt);
    if (sentDate < cutoff) {
      continue;
    }

    const type = String(data[i][typeIndex] || '');
    if (Object.prototype.hasOwnProperty.call(counts, type)) {
      counts[type] += 1;
    }
    counts.total += 1;
  }

  return counts;
}
