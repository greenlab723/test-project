function registerMember(formData) {
  const sheet = getMembersSheet();
  const header = sheet.getDataRange().getValues()[0] || [];
  ensureMemberColumns(header);

  const row = buildMemberRow(header, formData);
  sheet.appendRow(row);

  return {
    memberId: row[header.indexOf('member_id')],
    name: row[header.indexOf('name')],
    email: row[header.indexOf('email')],
  };
}

function buildMemberRow(header, formData) {
  const now = new Date();
  const memberId = generateMemberId();
  const values = new Array(header.length).fill('');

  setHeaderValue(header, values, 'member_id', memberId);
  setHeaderValue(header, values, 'name', formData.name || '');
  setHeaderValue(header, values, 'email', formData.email || '');
  setHeaderValue(header, values, 'status', 'pending');
  setHeaderValue(header, values, 'created_at', now);
  setHeaderValue(header, values, 'updated_at', now);
  setHeaderValue(header, values, 'form_data', JSON.stringify(formData));

  return values;
}

function activateMember(memberId) {
  const sheet = getMembersSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const idIndex = header.indexOf('member_id');
  const statusIndex = header.indexOf('status');
  const updatedIndex = header.indexOf('updated_at');

  if (idIndex === -1 || statusIndex === -1) {
    throw new Error('Members sheet must include member_id and status columns');
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]) === String(memberId)) {
      sheet.getRange(i + 1, statusIndex + 1).setValue('active');
      if (updatedIndex !== -1) {
        sheet.getRange(i + 1, updatedIndex + 1).setValue(new Date());
      }
      return;
    }
  }

  throw new Error('Member not found');
}

function findMemberByEmail(email) {
  const sheet = getMembersSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const emailIndex = header.indexOf('email');
  const idIndex = header.indexOf('member_id');
  const nameIndex = header.indexOf('name');
  const statusIndex = header.indexOf('status');

  if (emailIndex === -1) {
    throw new Error('Members sheet must include email column');
  }

  const normalized = String(email).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][emailIndex]).trim().toLowerCase() === normalized) {
      return {
        memberId: data[i][idIndex],
        name: data[i][nameIndex],
        email: data[i][emailIndex],
        status: data[i][statusIndex],
        rowIndex: i + 1,
      };
    }
  }

  return null;
}

function updateMember(memberId, updates, source) {
  const sheet = getMembersSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const idIndex = header.indexOf('member_id');
  const updatedIndex = header.indexOf('updated_at');
  const formDataIndex = header.indexOf('form_data');

  if (idIndex === -1) {
    throw new Error('Members sheet must include member_id column');
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]) === String(memberId)) {
      const changes = {};
      Object.keys(updates).forEach((key) => {
        const index = header.indexOf(key);
        if (index !== -1) {
          const previous = data[i][index];
          if (String(previous) !== String(updates[key])) {
            changes[key] = { before: previous, after: updates[key] };
          }
          sheet.getRange(i + 1, index + 1).setValue(updates[key]);
        }
      });
      if (formDataIndex !== -1) {
        sheet.getRange(i + 1, formDataIndex + 1).setValue(JSON.stringify(updates));
      }
      if (updatedIndex !== -1) {
        sheet.getRange(i + 1, updatedIndex + 1).setValue(new Date());
      }
      logMemberChange(memberId, source || 'member', changes);
      return;
    }
  }

  throw new Error('Member not found');
}

function updateMemberStatus(memberId, status) {
  const allowed = ['active', 'inactive', 'pending'];
  if (!allowed.includes(status)) {
    throw new Error('Invalid status');
  }
  updateMember(memberId, { status }, 'admin');
}

function logMemberChange(memberId, source, changes) {
  const sheet = getMemberChangeLogSheet();
  sheet.appendRow([
    new Date(),
    memberId,
    source || '',
    JSON.stringify(changes || {}),
  ]);
}

function setHeaderValue(header, row, key, value) {
  const index = header.indexOf(key);
  if (index !== -1) {
    row[index] = value;
  }
}

function buildFormData(e, fields) {
  const params = (e && e.parameter) || {};
  const data = {};
  fields.forEach((field) => {
    const value = params[field.key];
    if (field.required && !value) {
      throw new Error(`${field.label}は必須です。`);
    }
    data[field.key] = value || '';
  });

  return data;
}

function ensureMemberColumns(header) {
  const required = ['member_id', 'name', 'email', 'status', 'created_at', 'updated_at'];
  const missing = required.filter((key) => !header.includes(key));
  if (missing.length) {
    throw new Error(`Members sheet missing columns: ${missing.join(', ')}`);
  }
}
