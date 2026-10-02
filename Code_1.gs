const ANGEL_LOGO_ID = '1zMF8SrKMWcTMEteFhSDo8Ay6YgToKFl_';
const ROUND_LOGO_ID = '1W98lfH20PFsKw-K4x_Bc98cWyJIWTHpY';

const OWNER_EMAIL = 'angelpattasukadai08@gmail.com';
const ALEX_EMAIL = 'alexpriyan2021@gmail.com';
const EMAIL_RECIPIENTS = [ALEX_EMAIL, OWNER_EMAIL];


/* =========================================================
   WEBSITE
   ========================================================= */

const ONLINE_ESTIMATES_SHEET = 'ONLINE_ESTIMATES';
const ONLINE_API_KEY = 'JDUBIZnDzyU0gmZ_Hk8DXoJIfJihgMjz';

function doGet(e) {

  const params = (e && e.parameter) ? e.parameter : {};

  if (params.api === 'online_estimates') {
    return ContentService
      .createTextOutput(JSON.stringify(getOnlineEstimatesApi_(params.key, params.status)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Angel Pattasu Kadai')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents
      ? JSON.parse(e.postData.contents)
      : {};

    if (body.key !== ONLINE_API_KEY) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: 'Unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (body.action === 'mark_converted') {
      const result = markEstimateConverted_(body.estimateNo, body.billNo);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'Unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/* =========================================================
   GET PRODUCTS FROM GOOGLE SHEET
   ========================================================= */

function getProducts() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheets()[0];

  const data =
    sheet
      .getDataRange()
      .getValues();

  if (data.length < 2) {
    return [];
  }

  const products = [];

  for (let i = 0; i < data.length; i++) {

    const row = data[i];

    const sno =
      String(row[0] || '').trim();

    const name =
      String(row[1] || '').trim();

    const tamil =
      String(row[2] || '').trim();

    const rate =
      row[3] || '';

    const per =
      String(row[5] || '').trim();

    const category =
      String(row[6] || '').trim();

    // Column H = Image Link
    const imageLink =
      String(row[7] || '').trim();


    if (!name) {
      continue;
    }


    /* Skip header */

    const lowerName =
      name.toLowerCase();

    if (
      lowerName === 'name (english)' ||
      lowerName === 'name(english)'
    ) {
      continue;
    }


    /* Selling price */

    let priceText =
      String(row[4] || '');

    priceText =
      priceText
        .replace(/₹/g, '')
        .replace(/Rs\.?/gi, '')
        .replace(/,/g, '')
        .trim();


    const price =
      parseFloat(priceText);


    if (isNaN(price)) {
      continue;
    }


    products.push({

      sno: sno,
      name: name,
      tamil: tamil,
      rate: rate,
      price: price,
      per: per,

      category:
        category ||
        'OTHER PRODUCTS',

      // Convert Google Drive file links from Column H into browser image URLs.
      image: toProductImageUrl_(imageLink)

    });

  }


  return products;

}


/* =========================================================
   PRODUCT IMAGE URL HELPER
   Column H can contain a Google Drive file link or a direct image URL.
   ========================================================= */

function toProductImageUrl_(link) {

  const value = String(link || '').trim();

  if (!value) {
    return '';
  }

  // Google Drive file URL: /file/d/FILE_ID/view...
  let match = value.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);

  if (!match) {
    // Also support id=FILE_ID links.
    match = value.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  }

  if (match && match[1]) {
    return 'https://drive.google.com/thumbnail?id=' +
      encodeURIComponent(match[1]) + '&sz=w600';
  }

  return value;

}


/* =========================================================
   GET MAIN ANGEL LOGO
   ========================================================= */

function getAngelLogo() {

  try {

    const file =
      DriveApp.getFileById(
        ANGEL_LOGO_ID
      );

    const blob =
      file.getBlob();

    const bytes =
      blob.getBytes();

    const base64 =
      Utilities.base64Encode(
        bytes
      );


    return {

      success: true,

      mimeType:
        blob.getContentType(),

      data:
        base64

    };

  }

  catch (error) {

    return {

      success: false,

      error:
        String(error)

    };

  }

}


/* =========================================================
   GET ROUND LOGO
   FOR ESTIMATE PDF
   ========================================================= */

function getRoundLogo() {

  try {

    const file =
      DriveApp.getFileById(
        ROUND_LOGO_ID
      );

    const blob =
      file.getBlob();

    const bytes =
      blob.getBytes();

    const base64 =
      Utilities.base64Encode(
        bytes
      );


    return {

      success: true,

      mimeType:
        blob.getContentType(),

      data:
        base64

    };

  }

  catch (error) {

    return {

      success: false,

      error:
        String(error)

    };

  }

}


/* =========================================================
   CREATE ESTIMATE PDF
   ========================================================= */

function generateEstimatePdf(request) {

  if (!request || !request.items || !request.items.length) {
    throw new Error('No products found in the estimate.');
  }

  const customer = request.customer || {};
  const items = request.items || [];

  const billNo = getNextBillNumber();
  const dateText = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'Asia/Kolkata',
    'dd-MM-yyyy'
  );

  let originalValue = 0;
  let totalDiscount = 0;
  let subTotal = 0;
  let totalQty = 0;

  const pdfItems = items.map(function(item, index) {

    const qty = Math.max(0, Number(item.qty) || 0);
    const sellingPrice = Math.max(0, Number(item.price) || 0);
    const originalRate = Math.max(
      sellingPrice,
      parseMoney_(item.rate) || sellingPrice
    );

    const amount = sellingPrice * qty;
    const originalAmount = originalRate * qty;
    const discount = Math.max(0, originalAmount - amount);

    originalValue += originalAmount;
    totalDiscount += discount;
    subTotal += amount;
    totalQty += qty;

    const discountPercent =
      originalRate > 0
        ? (discount / originalAmount) * 100
        : 0;

    return {
      no: index + 1,
      product: safeText(item.name),
      tamil: safeText(item.tamil),
      qty: qty,
      // Price = original rate from Sheet (before discount)
      price: originalRate,
      discountPercent: discountPercent,
      discount: discount,
      // Dis Rate = final discounted selling rate
      rate: sellingPrice,
      amount: amount,
      per: safeText(item.per)
    };

  });

  const packingCharges = Math.max(
    0,
    Number(customer.packingCharges) || 0
  );

  const paidAmount = Math.max(
    0,
    Number(customer.paidAmount) || 0
  );

  const beforeRound = subTotal + packingCharges;

  const roundOff =
    Math.round(beforeRound) - beforeRound;

  const grandTotal =
    Math.round(beforeRound);

  const balance =
    Math.max(0, grandTotal - paidAmount);

  const roundLogo = getLogoDataUri_(ROUND_LOGO_ID);
  const angelLogo = getLogoDataUri_(ANGEL_LOGO_ID);

  const html = buildEstimateHtml_({
    billNo: billNo,
    dateText: dateText,
    customer: customer,
    items: pdfItems,
    totalQty: totalQty,
    originalValue: originalValue,
    totalDiscount: totalDiscount,
    subTotal: subTotal,
    packingCharges: packingCharges,
    roundOff: roundOff,
    grandTotal: grandTotal,
    paidAmount: paidAmount,
    balance: balance,
    angelLogo: angelLogo,
    roundLogo: roundLogo
  });

  const pdfBlob = HtmlService
    .createHtmlOutput(html)
    .getBlob()
    .setContentType(MimeType.HTML)
    .getAs(MimeType.PDF)
    .setName('Estimate-' + billNo + '.pdf');


  /* =========================================================
     SEND ESTIMATE EMAIL
     PDF IS SENT TO BOTH EMAIL ADDRESSES
     ========================================================= */

  const emailSubject =
    'ANGEL PATTASU KADAI - Estimate ' + billNo;

  const customerName =
    safeText(customer.name) || '-';

  const customerMobile =
    safeText(customer.mobile) || '-';

  const customerAddress =
    safeText(customer.address) || '-';

  const customerPayment =
    safeText(customer.payment) || 'Pending';

  const emailBody =
    'New estimate generated from ANGEL PATTASU KADAI website.\\n\\n' +
    'Bill No: ' + billNo + '\\n' +
    'Date: ' + dateText + '\\n' +
    'Customer: ' + customerName + '\\n' +
    'Mobile: ' + customerMobile + '\\n' +
    'Address: ' + customerAddress + '\\n' +
    'Payment: ' + customerPayment + '\\n' +
    'Total Items: ' + totalQty + '\\n' +
    'Original Value: ' + money_(originalValue) + '\\n' +
    'Total Discount: ' + money_(totalDiscount) + '\\n' +
    'Sub Total: ' + money_(subTotal) + '\\n' +
    'Packing Charges: ' + money_(packingCharges) + '\\n' +
    'Grand Total: ' + money_(grandTotal) + '\\n' +
    'Paid: ' + money_(paidAmount) + '\\n' +
    'Balance: ' + money_(balance) + '\\n\\n' +
    'The estimate PDF is attached to this email.\\n\\n' +
    'ANGEL PATTASU KADAI\\n' +
    '+91 82208 02867\\n' +
    'angelpattasukadai08@gmail.com';

  const emailHtmlBody =
    '<div style="font-family:Arial,sans-serif;color:#172033;line-height:1.6">' +
      '<h2 style="color:#08256f;margin-bottom:5px;">ANGEL PATTASU KADAI</h2>' +
      '<div style="color:#c18a18;font-weight:bold;margin-bottom:15px;">MOST RELIABLE BRAND</div>' +
      '<div style="border:1px solid #dce3f1;border-radius:8px;padding:15px;background:#f8faff;">' +
        '<h3 style="margin-top:0;color:#08256f;">New Estimate Generated</h3>' +
        '<p><b>Bill No:</b> ' + billNo + '<br>' +
        '<b>Date:</b> ' + dateText + '<br>' +
        '<b>Customer:</b> ' + customerName + '<br>' +
        '<b>Mobile:</b> ' + customerMobile + '<br>' +
        '<b>Address:</b> ' + customerAddress + '<br>' +
        '<b>Payment:</b> ' + customerPayment + '</p>' +
        '<hr style="border:none;border-top:1px solid #dce3f1;">' +
        '<p><b>Total Items:</b> ' + totalQty + '<br>' +
        '<b>Original Value:</b> ' + money_(originalValue) + '<br>' +
        '<b>Total Discount:</b> ' + money_(totalDiscount) + '<br>' +
        '<b>Sub Total:</b> ' + money_(subTotal) + '<br>' +
        '<b>Packing Charges:</b> ' + money_(packingCharges) + '<br>' +
        '<b>Grand Total:</b> <span style="font-size:18px;color:#d71920;font-weight:bold;">' + money_(grandTotal) + '</span><br>' +
        '<b>Paid:</b> ' + money_(paidAmount) + '<br>' +
        '<b>Balance:</b> ' + money_(balance) + '</p>' +
      '</div>' +
      '<p style="margin-top:15px;">The estimate PDF is attached to this email.</p>' +
      '<p style="color:#687286;font-size:12px;">ANGEL PATTASU KADAI<br>' +
      '+91 82208 02867<br>' +
      'angelpattasukadai08@gmail.com</p>' +
    '</div>';

  const emailErrors = [];

  EMAIL_RECIPIENTS.forEach(function(recipient) {

    try {

      MailApp.sendEmail({
        to: recipient,
        subject: emailSubject,
        body: emailBody,
        htmlBody: emailHtmlBody,
        attachments: [pdfBlob],
        name: 'ANGEL PATTASU KADAI'
      });

    } catch (error) {

      emailErrors.push(
        recipient + ': ' + String(error)
      );

    }

  });


  /* =========================================================
     SAVE ONLINE ESTIMATE TO GOOGLE SHEET
     ========================================================= */

  try {
    appendOnlineEstimate_({
      billNo: billNo,
      dateText: dateText,
      customerName: customerName,
      customerMobile: customerMobile,
      customerAddress: customerAddress,
      customerGstin: safeText(customer.gstin) || '-',
      customerPayment: customerPayment,
      paidAmount: paidAmount,
      packingCharges: packingCharges,
      originalValue: originalValue,
      totalDiscount: totalDiscount,
      subTotal: subTotal,
      grandTotal: grandTotal,
      balance: balance,
      items: pdfItems,
      pdfFileName: 'Estimate-' + billNo + '.pdf'
    });
  } catch (sheetError) {
    // Do not fail PDF generation if the online-estimate sheet cannot be written.
    console.error('ONLINE_ESTIMATES save failed: ' + String(sheetError));
  }

  return {
    success: true,
    fileName: 'Estimate-' + billNo + '.pdf',
    billNo: billNo,
    data: Utilities.base64Encode(pdfBlob.getBytes()),
    emailSent: emailErrors.length === 0,
    emailErrors: emailErrors
  };
}


/* =========================================================
   ONLINE ESTIMATES - WEBSITE -> DESKTOP BILLING APP
   ========================================================= */

function getOnlineEstimatesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ONLINE_ESTIMATES_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(ONLINE_ESTIMATES_SHEET);
    sheet.getRange(1, 1, 1, 18).setValues([[
      'Estimate No', 'Created At', 'Date', 'Customer', 'Mobile', 'Address',
      'GSTIN', 'Payment', 'Paid Amount', 'Packing Charges', 'Original Value',
      'Total Discount', 'Sub Total', 'Grand Total', 'Balance', 'Status',
      'Items JSON', 'PDF File'
    ]]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function appendOnlineEstimate_(record) {
  const sheet = getOnlineEstimatesSheet_();
  sheet.appendRow([
    record.billNo,
    new Date(),
    record.dateText,
    record.customerName,
    record.customerMobile,
    record.customerAddress,
    record.customerGstin,
    record.customerPayment,
    record.paidAmount,
    record.packingCharges,
    record.originalValue,
    record.totalDiscount,
    record.subTotal,
    record.grandTotal,
    record.balance,
    'NEW',
    JSON.stringify(record.items),
    record.pdfFileName
  ]);
}

function getOnlineEstimatesApi_(key, status) {
  if (key !== ONLINE_API_KEY) {
    return { success: false, error: 'Unauthorized' };
  }

  const sheet = getOnlineEstimatesSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[0]) continue;
    const rowStatus = String(r[15] || 'NEW').trim().toUpperCase();
    if (status && status !== 'ALL' && rowStatus !== String(status).toUpperCase()) continue;

    let items = [];
    try { items = JSON.parse(String(r[16] || '[]')); }
    catch (ignore) { items = []; }

    rows.push({
      estimateNo: String(r[0] || ''),
      createdAt: r[1] ? new Date(r[1]).toISOString() : '',
      date: String(r[2] || ''),
      customer: String(r[3] || ''),
      mobile: String(r[4] || ''),
      address: String(r[5] || ''),
      gstin: String(r[6] || ''),
      payment: String(r[7] || ''),
      paidAmount: Number(r[8] || 0),
      packingCharges: Number(r[9] || 0),
      originalValue: Number(r[10] || 0),
      totalDiscount: Number(r[11] || 0),
      subTotal: Number(r[12] || 0),
      grandTotal: Number(r[13] || 0),
      balance: Number(r[14] || 0),
      status: rowStatus,
      items: items,
      pdfFile: String(r[17] || '')
    });
  }

  rows.reverse();
  return { success: true, estimates: rows };
}

function markEstimateConverted_(estimateNo, billNo) {
  if (!estimateNo) return { success: false, error: 'Estimate number is required.' };

  const sheet = getOnlineEstimatesSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '') === String(estimateNo)) {
      sheet.getRange(i + 1, 16).setValue('CONVERTED');
      sheet.getRange(i + 1, 18).setValue(String(values[i][17] || '') + (billNo ? ' | Bill: ' + billNo : ''));
      return { success: true, estimateNo: String(estimateNo), billNo: String(billNo || '') };
    }
  }

  return { success: false, error: 'Estimate not found: ' + estimateNo };
}


/* =========================================================
   GET NEXT BILL NUMBER
   ========================================================= */

function getNextBillNumber() {

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    const properties =
      PropertiesService.getScriptProperties();

    // Keep online estimate numbering separate from any older ANG numbers.
    let number =
      parseInt(
        properties.getProperty('ANGON_BILL_NUMBER') || '0',
        10
      );

    if (isNaN(number)) {
      number = 0;
    }

    number++;

    properties.setProperty(
      'ANGON_BILL_NUMBER',
      String(number)
    );

    // Website / online estimate numbers use ANGON prefix.
    return 'ANGON-' +
      String(number).padStart(5, '0');

  }

  finally {

    lock.releaseLock();

  }

}


/* =========================================================
   LOGO DATA URI
   ========================================================= */

function getLogoDataUri_(fileId) {

  try {

    const file =
      DriveApp.getFileById(fileId);

    const blob =
      file.getBlob();

    return 'data:' +
      blob.getContentType() +
      ';base64,' +
      Utilities.base64Encode(
        blob.getBytes()
      );

  }

  catch (error) {

    return '';

  }

}


/* =========================================================
   BUILD ESTIMATE HTML
   ========================================================= */

function buildEstimateHtml_(data) {

  const customer = data.customer || {};

  const customerName =
    safeText(customer.name) || '-';

  const mobile =
    safeText(customer.mobile) || '-';

  const address =
    safeText(customer.address) || '-';

  const gstin =
    safeText(customer.gstin) || '-';

  const payment =
    safeText(customer.payment) || 'Pending';

  const rows = data.items.map(function(item) {

    const discountPercent =
      item.discountPercent > 0
        ? item.discountPercent.toFixed(0) + '%'
        : '0%';

    return `
      <tr>
        <td class="center">${item.no}</td>
        <td>
          <strong>${item.product}</strong>
          ${item.tamil ? `<div class="tamil">${item.tamil}</div>` : ''}
        </td>
        <td class="center">${item.qty}</td>
        <td class="right">${money_(item.price)}</td>
        <td class="center">${discountPercent}</td>
        <td class="right">${money_(item.discount)}</td>
        <td class="right">${money_(item.rate)}</td>
        <td class="right strong">${money_(item.amount)}</td>
      </tr>
    `;

  }).join('');

  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">

<style>

  @page {
    size: A4;
    margin: 12mm;
  }

  * {
    box-sizing: border-box;
  }

  body {
    margin: 0;
    font-family: Arial, Helvetica, sans-serif;
    color: #172033;
    font-size: 10px;
    background: white;
  }

  .page {
    width: 100%;
  }

  .top-line {
    height: 5px;
    background: linear-gradient(
      90deg,
      #08256f,
      #1655c9,
      #d71920
    );
    margin-bottom: 12px;
  }

  .header {
    display: table;
    width: 100%;
    border-bottom: 2px solid #d71920;
    padding-bottom: 10px;
  }

  .header-left,
  .header-right {
    display: table-cell;
    vertical-align: middle;
  }

  .header-left {
    width: 68%;
  }

  .header-right {
    width: 32%;
    text-align: right;
  }

  .angel-logo {
    width: 88px;
    height: 65px;
    object-fit: contain;
    vertical-align: middle;
    margin-right: 8px;
  }

  .brand {
    display: inline-block;
    vertical-align: middle;
  }

  .brand-name {
    color: #0b2b87;
    font-size: 21px;
    font-weight: 900;
    letter-spacing: .2px;
  }

  .tagline {
    color: #c18a18;
    font-size: 8px;
    font-weight: 800;
    letter-spacing: 2px;
    margin-top: 3px;
  }

  .address {
    color: #555f73;
    font-size: 8px;
    margin-top: 5px;
    line-height: 1.45;
  }

  .round-logo {
    width: 68px;
    height: 68px;
    object-fit: contain;
  }

  .estimate-title {
    margin-top: 12px;
    display: table;
    width: 100%;
  }

  .estimate-title-left,
  .estimate-title-right {
    display: table-cell;
    vertical-align: middle;
  }

  .estimate-title-left {
    width: 50%;
  }

  .estimate-title-right {
    width: 50%;
    text-align: right;
  }

  .estimate {
    display: inline-block;
    background: #08256f;
    color: white;
    padding: 7px 18px;
    border-radius: 5px;
    font-size: 16px;
    font-weight: 900;
    letter-spacing: 1px;
  }

  .meta {
    font-size: 9px;
    line-height: 1.7;
  }

  .meta strong {
    color: #0b2b87;
  }

  .customer-box {
    margin-top: 12px;
    border: 1px solid #dce3f1;
    border-radius: 7px;
    padding: 9px 10px;
    background: #f8faff;
  }

  .customer-grid {
    display: table;
    width: 100%;
  }

  .customer-col {
    display: table-cell;
    width: 50%;
    vertical-align: top;
    line-height: 1.65;
  }

  .label {
    color: #69748b;
    font-size: 8px;
    font-weight: 700;
    text-transform: uppercase;
  }

  .value {
    color: #172033;
    font-size: 9px;
    font-weight: 800;
  }

  table.items {
    width: 100%;
    border-collapse: collapse;
    margin-top: 13px;
  }

  .items th {
    background: #08256f;
    color: white;
    font-size: 8px;
    padding: 7px 5px;
    border: 1px solid #08256f;
  }

  .items td {
    border: 1px solid #dce2ed;
    padding: 6px 5px;
    font-size: 8px;
    vertical-align: middle;
  }

  .items tr:nth-child(even) td {
    background: #fafbfe;
  }

  .center {
    text-align: center;
  }

  .right {
    text-align: right;
  }

  .strong {
    font-weight: 900;
  }

  .tamil {
    color: #707b90;
    font-size: 7px;
    margin-top: 2px;
  }

  .summary-wrap {
    margin-top: 12px;
    display: table;
    width: 100%;
  }

  .summary-left,
  .summary-right {
    display: table-cell;
    vertical-align: top;
  }

  .summary-left {
    width: 54%;
    padding-right: 12px;
  }

  .summary-right {
    width: 46%;
  }

  .saving-box {
    border: 1px solid #dfe5f1;
    border-radius: 7px;
    padding: 9px 10px;
    background: #fbfcff;
  }

  .saving-title {
    color: #0b2b87;
    font-size: 9px;
    font-weight: 900;
    margin-bottom: 6px;
  }

  .saving-row {
    display: table;
    width: 100%;
    line-height: 1.8;
  }

  .saving-row span,
  .saving-row strong {
    display: table-cell;
  }

  .saving-row strong {
    text-align: right;
  }

  .note {
    margin-top: 9px;
    font-size: 7px;
    line-height: 1.45;
    color: #626d80;
  }

  .summary-table {
    width: 100%;
    border-collapse: collapse;
  }

  .summary-table td {
    padding: 4px 7px;
    border-bottom: 1px solid #e6eaf1;
    font-size: 8px;
  }

  .summary-table td:last-child {
    text-align: right;
    font-weight: 800;
  }

  .grand {
    background: #08256f;
    color: white;
    font-size: 11px !important;
    font-weight: 900;
  }

  .grand td {
    border: none !important;
    padding: 8px 7px !important;
  }

  .amount-words {
    margin-top: 10px;
    border: 1px solid #dce3f1;
    border-radius: 7px;
    padding: 8px 10px;
    background: #fffdf7;
  }

  .amount-words-title {
    color: #a66f0c;
    font-size: 7px;
    font-weight: 900;
    text-transform: uppercase;
  }

  .amount-words-value {
    margin-top: 3px;
    font-size: 9px;
    font-weight: 800;
  }

  .footer {
    margin-top: 14px;
    border-top: 1px solid #dfe4ed;
    padding-top: 8px;
    text-align: center;
    color: #687286;
    font-size: 7px;
    line-height: 1.5;
  }

  .thank-you {
    color: #0b2b87;
    font-size: 10px;
    font-weight: 900;
    margin-bottom: 3px;
  }

</style>
</head>

<body>

<div class="page">

  <div class="top-line"></div>

  <div class="header">

    <div class="header-left">

      ${data.angelLogo
        ? `<img class="angel-logo" src="${data.angelLogo}">`
        : ''
      }

      <div class="brand">

        <div class="brand-name">
          ANGEL PATTASU KADAI
        </div>

        <div class="tagline">
          MOST RELIABLE BRAND
        </div>

        <div class="address">
          AYYANAR NAGAR, NH-7, PATTAMPUDHUR<br>
          +91 82208 02867 &nbsp; | &nbsp;
          angelpattasukadai08@gmail.com
        </div>

      </div>

    </div>

    <div class="header-right">

      ${data.roundLogo
        ? `<img class="round-logo" src="${data.roundLogo}">`
        : ''
      }

    </div>

  </div>


  <div class="estimate-title">

    <div class="estimate-title-left">
      <div class="estimate">
        ESTIMATE
      </div>
    </div>

    <div class="estimate-title-right">

      <div class="meta">
        <strong>Bill No:</strong> ${data.billNo}<br>
        <strong>Date:</strong> ${data.dateText}
      </div>

    </div>

  </div>


  <div class="customer-box">

    <div class="customer-grid">

      <div class="customer-col">

        <div class="label">Customer</div>
        <div class="value">${customerName}</div>

        <div class="label">Mobile</div>
        <div class="value">${mobile}</div>

        <div class="label">Address</div>
        <div class="value">${address}</div>

      </div>

      <div class="customer-col">

        <div class="label">GSTIN</div>
        <div class="value">${gstin}</div>

        <div class="label">Payment</div>
        <div class="value">${payment}</div>

        <div class="label">Items / Quantity</div>
        <div class="value">${data.totalQty}</div>

      </div>

    </div>

  </div>


  <table class="items">

    <thead>

      <tr>
        <th style="width:5%">No.</th>
        <th style="width:29%">Product</th>
        <th style="width:7%">Qty</th>
        <th style="width:11%">Price</th>
        <th style="width:8%">Disc%</th>
        <th style="width:12%">Dis</th>
        <th style="width:12%">Dis Rate</th>
        <th style="width:16%">Amount</th>
      </tr>

    </thead>

    <tbody>

      ${rows}

    </tbody>

  </table>


  <div class="summary-wrap">

    <div class="summary-left">

      <div class="saving-box">

        <div class="saving-title">
          VALUE SUMMARY
        </div>

        <div class="saving-row">
          <span>Original Value</span>
          <strong>${money_(data.originalValue)}</strong>
        </div>

        <div class="saving-row">
          <span>Total Discount</span>
          <strong>${money_(data.totalDiscount)}</strong>
        </div>

        <div class="saving-row">
          <span>You Saved</span>
          <strong>${money_(data.totalDiscount)}</strong>
        </div>

      </div>

      <div class="amount-words">

        <div class="amount-words-title">
          Amount in Words
        </div>

        <div class="amount-words-value">
          ${numberToWords_(data.grandTotal)}
        </div>

      </div>

      <div class="note">
        I am a Composition Taxable Person, not eligible to collect tax on supplies.
      </div>

    </div>


    <div class="summary-right">

      <table class="summary-table">

        <tr>
          <td>Sub Total Value</td>
          <td>${money_(data.originalValue)}</td>
        </tr>

        <tr>
          <td>Total Discount</td>
          <td>${money_(data.totalDiscount)}</td>
        </tr>

        <tr>
          <td>Sub Total</td>
          <td>${money_(data.subTotal)}</td>
        </tr>

        <tr>
          <td>Packing Charges</td>
          <td>${money_(data.packingCharges)}</td>
        </tr>

        <tr>
          <td>Round Off</td>
          <td>${money_(data.roundOff)}</td>
        </tr>

        <tr class="grand">
          <td>GRAND TOTAL</td>
          <td>${money_(data.grandTotal)}</td>
        </tr>

        <tr>
          <td>Paid</td>
          <td>${money_(data.paidAmount)}</td>
        </tr>

        <tr>
          <td>Balance</td>
          <td>${money_(data.balance)}</td>
        </tr>

      </table>

    </div>

  </div>


  <div class="footer">

    <div class="thank-you">
      Thank You! Visit Again
    </div>

    ANGEL PATTASU KADAI &nbsp; | &nbsp;
    +91 82208 02867 &nbsp; | &nbsp;
    angelpattasukadai08@gmail.com

  </div>

</div>

</body>
</html>
`;

}


/* =========================================================
   PARSE MONEY
   ========================================================= */

function parseMoney_(value) {

  if (typeof value === 'number') {
    return isNaN(value) ? 0 : value;
  }

  const text = String(value == null ? '' : value)
    .replace(/₹/g, '')
    .replace(/Rs\.?/gi, '')
    .replace(/,/g, '')
    .trim();

  const number = parseFloat(text);

  return isNaN(number) ? 0 : number;

}


/* =========================================================
   MONEY
   ========================================================= */

function money_(value) {

  return '₹' +
    Number(value || 0).toLocaleString(
      'en-IN',
      {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      }
    );

}


/* =========================================================
   SAFE HTML TEXT
   ========================================================= */

function safeText(value) {

  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/\r?\n/g, '<br>');

}


/* =========================================================
   NUMBER TO WORDS
   ========================================================= */

function numberToWords_(number) {

  number = Math.round(Number(number) || 0);

  if (number === 0) {
    return 'Rupees Zero Only';
  }

  const ones = [
    '',
    'One',
    'Two',
    'Three',
    'Four',
    'Five',
    'Six',
    'Seven',
    'Eight',
    'Nine',
    'Ten',
    'Eleven',
    'Twelve',
    'Thirteen',
    'Fourteen',
    'Fifteen',
    'Sixteen',
    'Seventeen',
    'Eighteen',
    'Nineteen'
  ];

  const tens = [
    '',
    '',
    'Twenty',
    'Thirty',
    'Forty',
    'Fifty',
    'Sixty',
    'Seventy',
    'Eighty',
    'Ninety'
  ];

  function twoDigits(n) {

    if (n < 20) {
      return ones[n];
    }

    return tens[Math.floor(n / 10)] +
      (n % 10 ? ' ' + ones[n % 10] : '');

  }

  function threeDigits(n) {

    if (n < 100) {
      return twoDigits(n);
    }

    return ones[Math.floor(n / 100)] +
      ' Hundred' +
      (n % 100 ? ' ' + twoDigits(n % 100) : '');

  }

  let result = '';

  const crore = Math.floor(number / 10000000);
  number %= 10000000;

  const lakh = Math.floor(number / 100000);
  number %= 100000;

  const thousand = Math.floor(number / 1000);
  number %= 1000;

  const remainder = number;

  if (crore) {
    result += threeDigits(crore) + ' Crore ';
  }

  if (lakh) {
    result += twoDigits(lakh) + ' Lakh ';
  }

  if (thousand) {
    result += twoDigits(thousand) + ' Thousand ';
  }

  if (remainder) {
    result += threeDigits(remainder);
  }

  return 'Rupees ' +
    result.trim() +
    ' Only';

}
