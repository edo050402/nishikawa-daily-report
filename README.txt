NISHIKAWA AUTOMATION 業務日報 V2

Perubahan utama V2:
- 1 hari dapat memiliki banyak baris pekerjaan/JOB.
- Setiap baris memiliki JOB番号, 業務コード, 内容, 開始, 終了.
- JOB番号 dipilih dari database JOB yang dikelola Admin.
- JOB database menyimpan JOB番号, nama project, customer, dan default content.
- Sistem menolak waktu antarbaris yang tumpang tindih.
- Input waktu 15 menit.
- Jam kehadiran harian dipisahkan dari detail pekerjaan.
- Logo NISHIKAWA AUTOMATION ditampilkan lebih besar/jelas.
- Export Excel bulanan.
- Jika daily_report.db dari V1 diletakkan di folder ini, data V1 akan dimigrasikan otomatis saat pertama dijalankan.

Menjalankan:
python -m pip install -r requirements.txt
python app.py
Buka http://127.0.0.1:5000
