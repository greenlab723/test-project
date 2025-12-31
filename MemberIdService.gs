function generateMemberId() {
  const sheet = getMembersSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return buildMemberId(1);
  }

  const header = data[0];
  const idIndex = header.indexOf('member_id');
  if (idIndex === -1) {
    throw new Error('member_id column not found in Members sheet');
  }

  const currentYearSuffix = getCurrentYearSuffix();
  let maxSequence = 0;

  for (let i = 1; i < data.length; i++) {
    const rawId = data[i][idIndex];
    if (!rawId) {
      continue;
    }

    const idString = String(rawId).trim();
    if (idString.length < 6) {
      continue;
    }

    const yearSuffix = idString.slice(0, 2);
    const sequence = Number(idString.slice(2));
    if (yearSuffix === currentYearSuffix && Number.isFinite(sequence)) {
      maxSequence = Math.max(maxSequence, sequence);
    }
  }

  return buildMemberId(maxSequence + 1);
}

function buildMemberId(sequenceNumber) {
  const yearSuffix = getCurrentYearSuffix();
  const padded = padNumber(sequenceNumber, 4);
  return `${yearSuffix}${padded}`;
}

function getCurrentYearSuffix() {
  const year = new Date().getFullYear();
  return String(year).slice(-2);
}

function padNumber(value, length) {
  return String(value).padStart(length, '0');
}
