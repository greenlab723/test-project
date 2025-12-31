function createToken(memberId, type, payload) {
  const settings = getSettingsMap();
  const expireMinutes = Number(settings.TOKEN_EXPIRE_MINUTES) || 60;
  const token = Utilities.getUuid();
  const expiresAt = new Date(Date.now() + expireMinutes * 60 * 1000);

  const sheet = getTokensSheet();
  const header = sheet.getDataRange().getValues()[0] || [];
  ensureTokenColumns(header);
  const payloadIndex = header.indexOf('payload');

  const row = new Array(header.length).fill('');

  setHeaderValue(header, row, 'token', token);
  setHeaderValue(header, row, 'member_id', memberId);
  setHeaderValue(header, row, 'type', type);
  setHeaderValue(header, row, 'expires_at', expiresAt);
  setHeaderValue(header, row, 'created_at', new Date());
  if (payloadIndex !== -1 && payload) {
    row[payloadIndex] = JSON.stringify(payload);
  }

  sheet.appendRow(row);

  return {
    token,
    expiresAt,
  };
}

function verifyToken(token, type) {
  const sheet = getTokensSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const tokenIndex = header.indexOf('token');
  const memberIndex = header.indexOf('member_id');
  const typeIndex = header.indexOf('type');
  const expiresIndex = header.indexOf('expires_at');
  const usedIndex = header.indexOf('used_at');
  const payloadIndex = header.indexOf('payload');

  if ([tokenIndex, memberIndex, typeIndex, expiresIndex].includes(-1)) {
    throw new Error('Tokens sheet must include token, member_id, type, expires_at columns');
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][tokenIndex]) === String(token) && data[i][typeIndex] === type) {
      if (usedIndex !== -1 && data[i][usedIndex]) {
        return { valid: false, message: 'この認証リンクは既に使用されています。' };
      }
      const expiresAt = data[i][expiresIndex];
      if (expiresAt && new Date(expiresAt) < new Date()) {
        return { valid: false, message: 'この認証リンクは期限切れです。' };
      }
      let payload = null;
      if (payloadIndex !== -1 && data[i][payloadIndex]) {
        try {
          payload = JSON.parse(data[i][payloadIndex]);
        } catch (error) {
          payload = null;
        }
      }
      return { valid: true, memberId: data[i][memberIndex], payload };
    }
  }

  return { valid: false, message: '認証情報が見つかりません。' };
}

function markTokenUsed(token) {
  const sheet = getTokensSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const tokenIndex = header.indexOf('token');
  const usedIndex = header.indexOf('used_at');

  if (tokenIndex === -1 || usedIndex === -1) {
    return;
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][tokenIndex]) === String(token)) {
      sheet.getRange(i + 1, usedIndex + 1).setValue(new Date());
      return;
    }
  }
}

function ensureTokenColumns(header) {
  const required = ['token', 'member_id', 'type', 'expires_at', 'created_at'];
  const missing = required.filter((key) => !header.includes(key));
  if (missing.length) {
    throw new Error(`Tokens sheet missing columns: ${missing.join(', ')}`);
  }
}
