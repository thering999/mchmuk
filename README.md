# MCHMUK Analytics Dashboard

ระบบวิเคราะห์ข้อมูลเด็กรับธาตุเหล็กมาตรฐานกระทรวงสาธารณสุข (HDC Mode)

**Holographic Dashboard System v1.1**  
วิเคราะห์รายงานมาตรฐาน HDC: ร้อยละเด็กอายุ 6 เดือน – 5 ปี ได้รับยาน้ำเสริมธาตุเหล็ก (เป้าหมาย ≥ 75%)

## Features
- 📊 นำเข้าไฟล์ Excel / CSV ผ่าน Drag & Drop หรือ File Picker
- 🏥 MOPH HDC Mode: วิเคราะห์ KPI ครอบคลุม, HCT Lab, ภาวะซีด
- 📈 Charts: Area, Bar, Donut, Radar (ApexCharts)
- 🔍 ค้นหาและ filter ข้อมูลแบบ real-time
- 📤 Export CSV / JSON
- 🌟 Holographic UI Design (Glassmorphism + Neon)

## Tech Stack
- Vanilla HTML / CSS / JavaScript
- [SheetJS (xlsx)](https://sheetjs.com/) — อ่าน Excel
- [ApexCharts](https://apexcharts.com/) — Visualization
- [Lucide Icons](https://lucide.dev/)

## Live Demo
🔗 https://thering999.github.io/mchmuk/

## วิธีติดตั้ง
โปรเจกต์เป็น static site ไม่มีขั้นตอน build / ไม่มี package ให้ติดตั้ง  
ไลบรารี (SheetJS, ApexCharts, Lucide) โหลดผ่าน CDN ใน `index.html` จึงต้องต่ออินเทอร์เน็ต

**รันบนเครื่อง**
```bash
git clone https://github.com/thering999/mchmuk.git
cd mchmuk
# เปิดผ่าน web server ใดก็ได้ เช่น
python -m http.server 8000
# แล้วเข้า http://localhost:8000
```

**Deploy (GitHub Pages)**  
push เข้า branch `main` แล้ว workflow `.github/workflows/pages.yml` จะ deploy ทั้งโฟลเดอร์ขึ้น GitHub Pages อัตโนมัติ (สั่งรันเองได้ผ่าน `workflow_dispatch`)

**แหล่งข้อมูล**  
แอปดึงไฟล์ `tmp_exchange_data.xlsx` และ `tmp_exchange_data_meta.json` จาก branch `main` ของ repo `thering999/mchmuk` (ตั้งค่าที่ `GITHUB_CONFIG` ใน `app.js`) — ถ้า fork ไปใช้ ต้องแก้ `owner` / `repo` / `branch` ให้ตรง

## Usage
เปิด `index.html` ผ่าน web server หรือดูผ่าน GitHub Pages  
จากนั้น upload ไฟล์ Excel จาก HDC รายงานธาตุเหล็กเด็ก

---
© 2026 MCHMUK Dev Space
