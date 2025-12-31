function getMembersForAdmin(options) {
  const sheet = getMembersSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const nameIndex = header.indexOf('name');
  const emailIndex = header.indexOf('email');
  const statusIndex = header.indexOf('status');

  if ([nameIndex, emailIndex, statusIndex].includes(-1)) {
    throw new Error('Members sheet must include name, email, status columns');
  }

  const query = (options.query || '').toLowerCase();
  const statusFilter = options.status || '';
  const createdFrom = parseDateFilter(options.createdFrom);
  const createdTo = parseDateFilter(options.createdTo, true);
  const updatedFrom = parseDateFilter(options.updatedFrom);
  const updatedTo = parseDateFilter(options.updatedTo, true);
  const limit = options.limit || 50;
  const page = options.page || 1;
  const offset = (page - 1) * limit;
  const members = [];
  let matchCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[nameIndex] || '';
    const email = row[emailIndex] || '';
    const status = row[statusIndex] || '';
    const createdAt = row[header.indexOf('created_at')];
    const updatedAt = row[header.indexOf('updated_at')];

    if (statusFilter && String(status) !== statusFilter) {
      continue;
    }

    if (createdFrom && (!createdAt || new Date(createdAt) < createdFrom)) {
      continue;
    }

    if (createdTo && (!createdAt || new Date(createdAt) > createdTo)) {
      continue;
    }

    if (updatedFrom && (!updatedAt || new Date(updatedAt) < updatedFrom)) {
      continue;
    }

    if (updatedTo && (!updatedAt || new Date(updatedAt) > updatedTo)) {
      continue;
    }

    if (query) {
      const haystack = `${name} ${email}`.toLowerCase();
      if (!haystack.includes(query)) {
        continue;
      }
    }

    if (matchCount >= offset && members.length < limit) {
      members.push({
        memberId: row[header.indexOf('member_id')],
        name,
        email,
        status,
        createdAt,
        updatedAt,
      });
    }

    matchCount += 1;
  }

  return { members, total: matchCount };
}

function parseDateFilter(value, endOfDay) {
  if (!value) {
    return null;
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return null;
  }
  if (endOfDay) {
    date.setHours(23, 59, 59, 999);
  } else {
    date.setHours(0, 0, 0, 0);
  }
  return date;
}

function getMemberRowsForExport(options) {
  const sheet = getMembersSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const idIndex = header.indexOf('member_id');
  const nameIndex = header.indexOf('name');
  const emailIndex = header.indexOf('email');
  const statusIndex = header.indexOf('status');
  const createdIndex = header.indexOf('created_at');
  const updatedIndex = header.indexOf('updated_at');

  if ([idIndex, nameIndex, emailIndex, statusIndex].includes(-1)) {
    throw new Error('Members sheet must include member_id, name, email, status columns');
  }

  const query = (options.query || '').toLowerCase();
  const statusFilter = options.status || '';
  const createdFrom = parseDateFilter(options.createdFrom);
  const createdTo = parseDateFilter(options.createdTo, true);
  const updatedFrom = parseDateFilter(options.updatedFrom);
  const updatedTo = parseDateFilter(options.updatedTo, true);
  const fieldKeys = options.fieldKeys || [];
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[nameIndex] || '';
    const email = row[emailIndex] || '';
    const status = row[statusIndex] || '';
    const createdAt = row[createdIndex];
    const updatedAt = row[updatedIndex];

    if (statusFilter && String(status) !== statusFilter) {
      continue;
    }

    if (createdFrom && (!createdAt || new Date(createdAt) < createdFrom)) {
      continue;
    }

    if (createdTo && (!createdAt || new Date(createdAt) > createdTo)) {
      continue;
    }

    if (updatedFrom && (!updatedAt || new Date(updatedAt) < updatedFrom)) {
      continue;
    }

    if (updatedTo && (!updatedAt || new Date(updatedAt) > updatedTo)) {
      continue;
    }

    if (query) {
      const haystack = `${name} ${email}`.toLowerCase();
      if (!haystack.includes(query)) {
        continue;
      }
    }

    const formData = parseFormData(row[header.indexOf('form_data')]);
    const extraValues = fieldKeys.map((key) => (formData ? formData[key] || '' : ''));
    rows.push([
      row[idIndex],
      name,
      email,
      status,
      createdAt || '',
      updatedAt || '',
      ...extraValues,
    ]);
  }

  return rows;
}

function parseFormData(value) {
  if (!value) {
    return null;
  }
  try {
    return JSON.parse(value);
  } catch (error) {
    return null;
  }
}
