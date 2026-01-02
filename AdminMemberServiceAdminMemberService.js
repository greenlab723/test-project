function getMemberDetailForAdmin(memberId) {
  const sheet = getMembersSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const idIndex = header.indexOf('member_id');
  const nameIndex = header.indexOf('name');
  const emailIndex = header.indexOf('email');
  const statusIndex = header.indexOf('status');
  const createdIndex = header.indexOf('created_at');
  const updatedIndex = header.indexOf('updated_at');
  const formDataIndex = header.indexOf('form_data');

  if (idIndex === -1) {
    throw new Error('Members sheet must include member_id column');
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]) === String(memberId)) {
      const rawFormData = data[i][formDataIndex];
      let formData = null;
      if (rawFormData) {
        try {
          formData = JSON.parse(rawFormData);
        } catch (error) {
          formData = null;
        }
      }

      return {
        memberId: data[i][idIndex],
        name: data[i][nameIndex],
        email: data[i][emailIndex],
        status: data[i][statusIndex],
        createdAt: data[i][createdIndex],
        updatedAt: data[i][updatedIndex],
        formData: formData,
      };
    }
  }

  throw new Error('Member not found');
}
function getMemberDetailForAdmin(memberId) {
  const sheet = getMembersSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const idIndex = header.indexOf('member_id');
  const nameIndex = header.indexOf('name');
  const emailIndex = header.indexOf('email');
  const statusIndex = header.indexOf('status');
  const createdIndex = header.indexOf('created_at');
  const updatedIndex = header.indexOf('updated_at');
  const formDataIndex = header.indexOf('form_data');

  if (idIndex === -1) {
    throw new Error('Members sheet must include member_id column');
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]) === String(memberId)) {
      const rawFormData = data[i][formDataIndex];
      let formData = null;
      if (rawFormData) {
        try {
          formData = JSON.parse(rawFormData);
        } catch (error) {
          formData = null;
        }
      }

      return {
        memberId: data[i][idIndex],
        name: data[i][nameIndex],
        email: data[i][emailIndex],
        status: data[i][statusIndex],
        createdAt: data[i][createdIndex],
        updatedAt: data[i][updatedIndex],
        formData: formData,
      };
    }
  }

  throw new Error('Member not found');
}
