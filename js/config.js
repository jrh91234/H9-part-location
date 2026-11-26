// URL ของ Apps Script Web App (/exec) สำหรับเก็บการเปลี่ยนคู่ Part ↔ Location (เมนู ⚙ ตั้งค่า)
// GitHub Actions อัปเดต deployment เดิมทุกครั้งที่ push โฟลเดอร์ appscript/ → URL นี้จึงไม่เปลี่ยน
// ถ้ายังว่าง: เมนูตั้งค่าจะเก็บการเปลี่ยนไว้ในเครื่องนั้นเท่านั้น (ใส่ URL ชั่วคราวได้ในหน้า ⚙ ตั้งค่า → ขั้นสูง)
window.H9_CONFIG = {
  API_URL: 'https://script.google.com/macros/s/AKfycbw2uGiJeKOy1Lgt8Utupf_LnuMQ70qo5I-XLG3jJk_U477quGu1FSh5JTpka9sJSii1Hg/exec'
};
