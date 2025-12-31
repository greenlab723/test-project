function getMemberChangeLogForAdmin(options) {
  const sheet = getMemberChangeLogSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const changedIndex = header.indexOf('changed_at');
  const memberIndex = header.indexOf('member_id');
  const sourceIndex = header.indexOf('source');
  const changesIndex = header.indexOf('changes');

  if ([changedIndex, memberIndex, sourceIndex, changesIndex].includes(-1)) {
    throw new Error('MemberChangeLog sheet must include changed_at, member_id, source, changes');
  }

  const query = (options.query || '').toLowerCase();
  const memberFilter = options.memberId || '';
  const limit = options.limit || 50;
  const page = options.page || 1;
  const offset = (page - 1) * limit;
  const logs = [];
  let matchCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const memberId = row[memberIndex] || '';
    const source = row[sourceIndex] || '';
    const changes = row[changesIndex] || '';

    if (memberFilter && String(memberId) !== String(memberFilter)) {
      continue;
    }

    if (query) {
      const haystack = `${memberId} ${source} ${changes}`.toLowerCase();
      if (!haystack.includes(query)) {
        continue;
      }
    }

    if (matchCount >= offset && logs.length < limit) {
      logs.push({
        changedAt: row[changedIndex],
        memberId,
        source,
        changes,
      });
    }

    matchCount += 1;
  }

  return { logs, total: matchCount };
}

function getMemberChangeLogRowsForExport(options) {
  const sheet = getMemberChangeLogSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const changedIndex = header.indexOf('changed_at');
  const memberIndex = header.indexOf('member_id');
  const sourceIndex = header.indexOf('source');
  const changesIndex = header.indexOf('changes');

  if ([changedIndex, memberIndex, sourceIndex, changesIndex].includes(-1)) {
    throw new Error('MemberChangeLog sheet must include changed_at, member_id, source, changes');
  }

  const query = (options.query || '').toLowerCase();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const memberId = row[memberIndex] || '';
    const source = row[sourceIndex] || '';
    const changes = row[changesIndex] || '';

    if (query) {
      const haystack = `${memberId} ${source} ${changes}`.toLowerCase();
      if (!haystack.includes(query)) {
        continue;
      }
    }

    rows.push([row[changedIndex], memberId, source, changes]);
  }

  return rows;
}
