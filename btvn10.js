class Student {
    constructor(stt, msv, fullName) {
        this.stt = stt;
        this.msv = String(msv).trim();
        this.fullName = String(fullName).trim();
        // Tự động lấy Khóa học từ 2 số đầu của MSV (VD: 27A... -> Khóa 27)
        this.khoaHoc = `Khóa ${this.msv.substring(0, 2)}`;
        // Phân loại khoa dựa trên mã ngành trong MSV
        this.tenKhoa = this.msv.includes('40429') ? "CNTT" : "HTTTQL";
        this.email = this.generateEmail();
    }

    generateEmail() {
        // Xử lý chuẩn hóa tên: xóa dấu, viết thường, tách mảng
        let parts = this.fullName
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/đ/g, 'd')
            .replace(/Đ/g, 'D')
            .toLowerCase()
            .split(/\s+/); // split theo 1 hoặc nhiều khoảng trắng

        if (parts.length > 0) {
            let ten = parts.pop(); // Lấy tên chính
            let initials = parts.map(p => p[0]).join(''); // Lấy chữ cái đầu họ đệm
            return `${ten}${initials}.${this.msv.toLowerCase()}@hvnh.edu.vn`;
        }
        return "";
    }
}

document.getElementById('excelFile').addEventListener('change', e => {
    let file = e.target.files[0];
    if (!file) return;

    let reader = new FileReader();
    reader.onload = ev => {
        let data = new Uint8Array(ev.target.result);
        let workbook = XLSX.read(data, { type: 'array' });
        
        // Đọc dữ liệu dạng mảng (header: 1 giúp lấy theo chỉ số cột 0, 1, 2...)
        let rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
        
        let html = '';
        // Bắt đầu từ i = 1 để bỏ qua dòng tiêu đề của Excel
        for (let i = 1; i < rows.length; i++) {
            let stt = rows[i][0];
            let msv = rows[i][1];
            let hoTen = rows[i][2];

            // Kiểm tra nếu dòng có đủ MSV và Họ Tên thì mới tạo đối tượng
            if (msv && hoTen) {
                let s = new Student(stt || i, msv, hoTen);
                html += `
                    <tr>
                        <td>${s.stt}</td>
                        <td>${s.msv}</td>
                        <td>${s.fullName}</td>
                        <td>${s.khoaHoc}</td>
                        <td>${s.tenKhoa}</td>
                        <td><a href="mailto:${s.email}">${s.email}</a></td>
                    </tr>`;
            }
        }
        document.querySelector('#studentTable tbody').innerHTML = html;
    };
    reader.readAsArrayBuffer(file);
});