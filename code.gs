function myFunction() {
	Logger.log(MailApp.getRemainingDailyQuota());
	const rows = [94, 95, 96, 120];

	for (let i = 0; i < rows.length; i++) {
		sendEmailByRow(rows[i]);
	}
}

function sendEmailByRow(row) {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses");
	if (!sheet) throw new Error('Response sheet not found');

	row = Number(row);
	if (!row || row < 2) throw new Error('Invalid row number');

	try {
		if (MailApp.getRemainingDailyQuota() < 1) throw new Error('No email quota remaining today');

		const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
		const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

		const record = {};
		headers.forEach((h, i) => record[h] = values[i]);

		const membershipID = String(values[2] || '').trim();
		if (!membershipID) throw new Error('No membership ID in column C');

		const emailAddress = String(record['Email Address'] || record['ที่อยู่อีเมล'] || '').trim();
		if (!emailAddress) throw new Error('Missing email address');

		const emailData = {
			membershipID: membershipID,
			parentFirstName: record["ชื่อผู้ปกครอง"] || '',
			parentLastName: record["นามสกุลผู้ปกครอง"] || '',
			studentFirstName: record["ชื่อนักเรียน"] || '',
			studentLastName: record["นามสกุลนักเรียน"] || '',
			jerseySize: record["ไซส์เสื้อ Jersey"] || '-',
			jerseyText: record["ตัวอักษรบนเสื้อ Jersey"] || '-',
			jerseyNumber: record["ตัวเลขบนเสื้อ Jersey"] || '-',
			poloSize: record["ไซส์เสื้อโปโล"] || '-',
			poloGender: record["เพศเสื้อโปโล"] || '-'
		};

		const subject = 'ขอบคุณสำหรับการกรอกแบบฟอร์มสมัครสมาชิกและจองไซส์เสื้อ Triamudom Family';
		const plainText = generatePlainText(emailData);
		const htmlBody = generateHtmlBody(emailData);

		GmailApp.sendEmail(emailAddress, subject, plainText, {
			htmlBody: htmlBody,
			name: 'ข้อมูลการสมัครสมาชิก',
			noReply: true
		});

		sheet.getRange(row, 32).setValue("TRUE");

		sheet
			.getRange(row, 1, 1, sheet.getLastColumn())
			.setFontFamily('Anuphan');

		const monoRanges = [
			'A' + row, 'B' + row, 'C' + row, 'H' + row, 'J' + row,
			'Q' + row, 'R' + row, 'S' + row, 'T' + row, 'AD' + row
		];
		sheet.getRangeList(monoRanges).setFontFamily('JetBrains Mono');

		sheet.getRange('A' + row).setHorizontalAlignment('right');
		sheet.getRangeList(['B' + row, 'AD' + row]).setHorizontalAlignment('left');
		sheet.getRangeList([
			'C' + row, 'H' + row, 'I' + row, 'J' + row, 'K' + row,
			'L' + row, 'M' + row, 'Q' + row, 'R' + row, 'S' + row,
			'T' + row, 'U' + row, 'AE' + row
		]).setHorizontalAlignment('center');

		sheet.getRange(row, 33).setValue('Retry Successfully');
		Logger.log('Retry succeeded for row %s', row);
	} catch (err) {
		sheet.getRange(row, 33).setValue('Retry Failed: ' + err.message);
		throw err;
	}
}

function generatePlainText(data) {
	return `
	`.trim();
}

function generateHtmlBody(data) {
	return `
	`.trim();
}

function escapeHtml_(text) {
	return String(text ?? '')
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&#39;');
}