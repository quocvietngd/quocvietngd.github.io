import "./style.css";
import { Chart, registerables } from "chart.js";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";

Chart.register(...registerables);

const DEPARTMENTS = ["Nhân sự", "Kinh doanh", "Marketing", "Kỹ thuật", "Vận hành", "Tài chính"];
const DEPARTMENT_MONTHLY_REVENUE_TARGET = {
  "Nhân sự": 220000000,
  "Kinh doanh": 920000000,
  "Marketing": 410000000,
  "Kỹ thuật": 560000000,
  "Vận hành": 360000000,
  "Tài chính": 300000000
};
const CUSTOMER_STATUSES = ["Đã gọi", "Không nghe máy", "Đang cân nhắc", "Đã ký", "Đang chăm lại"];
const CUSTOMER_SOURCES = ["Nhập tay", "Website", "Facebook", "Zalo", "Giới thiệu", "API", "Google Sheets"];
/* eslint-disable */
const VN_ADDRESSES = {
  "TP. Hồ Chí Minh": {
    "Quận 1": ["P. Bến Nghé","P. Bến Thành","P. Cầu Kho","P. Cầu Ông Lãnh","P. Cô Giang","P. Đa Kao","P. Nguyễn Cư Trinh","P. Nguyễn Thái Bình","P. Phạm Ngũ Lão","P. Tân Định"],
    "Quận 3": ["P. 1","P. 2","P. 3","P. 4","P. 5","P. 6","P. 7","P. 8","P. 9","P. 10","P. 11","P. 12","P. 13","P. 14"],
    "Quận 4": ["P. 1","P. 2","P. 3","P. 4","P. 6","P. 8","P. 9","P. 10","P. 13","P. 14","P. 15","P. 16","P. 18"],
    "Quận 5": ["P. 1","P. 2","P. 3","P. 4","P. 5","P. 6","P. 7","P. 8","P. 9","P. 10","P. 11","P. 12","P. 13","P. 14","P. 15"],
    "Quận 6": ["P. 1","P. 2","P. 3","P. 4","P. 5","P. 6","P. 7","P. 8","P. 9","P. 10","P. 11","P. 12","P. 13","P. 14"],
    "Quận 7": ["P. Tân Thuận Đông","P. Tân Thuận Tây","P. Tân Kiểng","P. Tân Hưng","P. Bình Thuận","P. Tân Quy","P. Phú Thuận","P. Tân Phú","P. Tân Phong","P. Phú Mỹ"],
    "Quận 8": ["P. 1","P. 2","P. 3","P. 4","P. 5","P. 6","P. 7","P. 8","P. 9","P. 10","P. 11","P. 12","P. 13","P. 14","P. 15","P. 16"],
    "Quận 10": ["P. 1","P. 2","P. 3","P. 4","P. 5","P. 6","P. 7","P. 8","P. 9","P. 10","P. 11","P. 12","P. 13","P. 14","P. 15"],
    "Quận 11": ["P. 1","P. 2","P. 3","P. 4","P. 5","P. 6","P. 7","P. 8","P. 9","P. 10","P. 11","P. 12","P. 13","P. 14","P. 15","P. 16"],
    "Quận 12": ["P. An Phú Đông","P. Đông Hưng Thuận","P. Hiệp Thành","P. Tân Chánh Hiệp","P. Tân Hưng Thuận","P. Tân Thới Hiệp","P. Tân Thới Nhất","P. Thạnh Lộc","P. Thạnh Xuân","P. Thới An","P. Trung Mỹ Tây"],
    "Quận Bình Thạnh": ["P. 1","P. 2","P. 3","P. 5","P. 6","P. 7","P. 11","P. 12","P. 13","P. 14","P. 15","P. 17","P. 19","P. 21","P. 22","P. 24","P. 25","P. 26","P. 27","P. 28"],
    "Quận Tân Bình": ["P. 1","P. 2","P. 3","P. 4","P. 5","P. 6","P. 7","P. 8","P. 9","P. 10","P. 11","P. 12","P. 13","P. 14","P. 15"],
    "Quận Tân Phú": ["P. Hiệp Tân","P. Hòa Thạnh","P. Phú Thạnh","P. Phú Thọ Hòa","P. Phú Trung","P. Sơn Kỳ","P. Tân Quý","P. Tân Sơn Nhì","P. Tân Thành","P. Tân Thới Hòa","P. Tây Thạnh"],
    "Quận Gò Vấp": ["P. 1","P. 3","P. 4","P. 5","P. 6","P. 7","P. 8","P. 9","P. 10","P. 11","P. 12","P. 14","P. 15","P. 16","P. 17"],
    "Quận Phú Nhuận": ["P. 1","P. 2","P. 3","P. 4","P. 5","P. 7","P. 8","P. 9","P. 10","P. 11","P. 13","P. 14","P. 15","P. 17"],
    "TP. Thủ Đức": ["P. An Khánh","P. An Lợi Đông","P. An Phú","P. Bình Chiểu","P. Bình Thọ","P. Bình Trưng Đông","P. Bình Trưng Tây","P. Cát Lái","P. Hiệp Bình Chánh","P. Hiệp Bình Phước","P. Hiệp Phú","P. Linh Chiểu","P. Linh Đông","P. Linh Tây","P. Linh Trung","P. Linh Xuân","P. Long Bình","P. Long Phước","P. Long Thạnh Mỹ","P. Long Trường","P. Phú Hữu","P. Phước Bình","P. Phước Long A","P. Phước Long B","P. Tân Phú","P. Tam Bình","P. Tam Phú","P. Tăng Nhơn Phú A","P. Tăng Nhơn Phú B","P. Thảo Điền","P. Thủ Thiêm","P. Trường Thạnh","P. Trường Thọ"],
    "H. Bình Chánh": ["TT. Tân Túc","Xã Bình Chánh","Xã Bình Hưng","Xã Bình Lợi","Xã Đa Phước","Xã Hưng Long","Xã Lê Minh Xuân","Xã Phạm Văn Hai","Xã Phong Phú","Xã Qui Đức","Xã Tân Kiên","Xã Tân Nhựt","Xã Tân Quý Tây","Xã Vĩnh Lộc A","Xã Vĩnh Lộc B"],
    "H. Nhà Bè": ["TT. Nhà Bè","Xã Hiệp Phước","Xã Long Thới","Xã Nhơn Đức","Xã Phú Xuân","Xã Phước Kiển","Xã Phước Lộc"],
    "H. Hóc Môn": ["TT. Hóc Môn","Xã Bà Điểm","Xã Đông Thạnh","Xã Nhị Bình","Xã Tân Hiệp","Xã Tân Thới Nhì","Xã Tân Xuân","Xã Thới Tam Thôn","Xã Trung Chánh","Xã Xuân Thới Đông","Xã Xuân Thới Sơn","Xã Xuân Thới Thượng"],
    "H. Củ Chi": ["TT. Củ Chi","Xã An Nhơn Tây","Xã An Phú","Xã Bình Mỹ","Xã Hòa Phú","Xã Nhuận Đức","Xã Phạm Văn Cội","Xã Phú Hòa Đông","Xã Phú Mỹ Hưng","Xã Phước Hiệp","Xã Phước Thạnh","Xã Phước Vĩnh An","Xã Tân An Hội","Xã Tân Phú Trung","Xã Tân Thạnh Đông","Xã Tân Thạnh Tây","Xã Tân Thông Hội","Xã Thái Mỹ","Xã Trung An","Xã Trung Lập Hạ","Xã Trung Lập Thượng"],
    "H. Cần Giờ": ["TT. Cần Thạnh","Xã An Thới Đông","Xã Bình Khánh","Xã Long Hòa","Xã Lý Nhơn","Xã Tam Thôn Hiệp","Xã Thạnh An"]
  },
  "Hà Nội": {
    "Q. Ba Đình": ["P. Cống Vị","P. Điện Biên","P. Đội Cấn","P. Giảng Võ","P. Kim Mã","P. Liễu Giai","P. Ngọc Hà","P. Ngọc Khánh","P. Nguyễn Trung Trực","P. Phúc Xá","P. Quán Thánh","P. Thành Công","P. Trúc Bạch","P. Vĩnh Phúc"],
    "Q. Hoàn Kiếm": ["P. Chương Dương","P. Cửa Đông","P. Cửa Nam","P. Đồng Xuân","P. Hàng Bạc","P. Hàng Bài","P. Hàng Bồ","P. Hàng Buồm","P. Hàng Gai","P. Hàng Mã","P. Hàng Trống","P. Lý Thái Tổ","P. Phan Chu Trinh","P. Phúc Tân","P. Tràng Tiền","P. Trần Hưng Đạo"],
    "Q. Đống Đa": ["P. Cát Linh","P. Hàng Bột","P. Khâm Thiên","P. Kim Liên","P. Láng Hạ","P. Láng Thượng","P. Nam Đồng","P. Ngã Tư Sở","P. Ô Chợ Dừa","P. Phương Liên","P. Phương Mai","P. Quang Trung","P. Quốc Tử Giám","P. Thịnh Quang","P. Thổ Quan","P. Trung Liệt","P. Trung Phụng","P. Trung Tự","P. Văn Chương","P. Văn Miếu","P. Xã Đàn"],
    "Q. Hai Bà Trưng": ["P. Bách Khoa","P. Bạch Đằng","P. Bạch Mai","P. Cầu Dền","P. Đồng Nhân","P. Đồng Tâm","P. Lê Đại Hành","P. Minh Khai","P. Nguyễn Du","P. Phạm Đình Hổ","P. Phố Huế","P. Quỳnh Lôi","P. Quỳnh Mai","P. Thanh Lương","P. Thanh Nhàn","P. Trương Định","P. Vĩnh Tuy"],
    "Q. Hoàng Mai": ["P. Đại Kim","P. Định Công","P. Giáp Bát","P. Hoàng Liệt","P. Hoàng Văn Thụ","P. Lĩnh Nam","P. Mai Động","P. Tân Mai","P. Thanh Trì","P. Thịnh Liệt","P. Trần Phú","P. Tương Mai","P. Vĩnh Hưng","P. Yên Sở"],
    "Q. Thanh Xuân": ["P. Hạ Đình","P. Khương Đình","P. Khương Mai","P. Khương Trung","P. Kim Giang","P. Nhân Chính","P. Phương Liệt","P. Thanh Xuân Bắc","P. Thanh Xuân Nam","P. Thanh Xuân Trung","P. Thượng Đình"],
    "Q. Cầu Giấy": ["P. Dịch Vọng","P. Dịch Vọng Hậu","P. Mai Dịch","P. Nghĩa Đô","P. Nghĩa Tân","P. Quan Hoa","P. Trung Hòa","P. Yên Hòa"],
    "Q. Nam Từ Liêm": ["P. Cầu Diễn","P. Đại Mỗ","P. Mễ Trì","P. Mỹ Đình 1","P. Mỹ Đình 2","P. Phú Đô","P. Phương Canh","P. Tây Mỗ","P. Trung Văn","P. Xuân Phương"],
    "Q. Bắc Từ Liêm": ["P. Cổ Nhuế 1","P. Cổ Nhuế 2","P. Đông Ngạc","P. Đức Thắng","P. Liên Mạc","P. Minh Khai","P. Phú Diễn","P. Phúc Diễn","P. Tây Tựu","P. Thượng Cát","P. Thụy Phương","P. Xuân Đỉnh","P. Xuân Tảo"],
    "Q. Long Biên": ["P. Bồ Đề","P. Cự Khối","P. Đức Giang","P. Gia Thụy","P. Giang Biên","P. Long Biên","P. Ngọc Lâm","P. Ngọc Thụy","P. Phúc Đồng","P. Phúc Lợi","P. Sài Đồng","P. Thạch Bàn","P. Thượng Thanh","P. Việt Hưng"],
    "Q. Tây Hồ": ["P. Bưởi","P. Nhật Tân","P. Phú Thượng","P. Quảng An","P. Tứ Liên","P. Xuân La","P. Yên Phụ","P. Thụy Khuê"],
    "Q. Hà Đông": ["P. Biên Giang","P. Dương Nội","P. Đồng Mai","P. Hà Cầu","P. Kiến Hưng","P. La Khê","P. Mộ Lao","P. Nguyễn Trãi","P. Phú La","P. Phú Lãm","P. Phú Lương","P. Phúc La","P. Quang Trung","P. Vạn Phúc","P. Văn Quán","P. Yên Nghĩa","P. Yên Sở"],
    "TX. Sơn Tây": [],"H. Ba Vì": [],"H. Chương Mỹ": [],"H. Đan Phượng": [],"H. Đông Anh": [],"H. Gia Lâm": [],"H. Hoài Đức": [],"H. Mê Linh": [],"H. Mỹ Đức": [],"H. Phú Xuyên": [],"H. Phúc Thọ": [],"H. Quốc Oai": [],"H. Sóc Sơn": [],"H. Thạch Thất": [],"H. Thanh Oai": [],"H. Thanh Trì": [],"H. Thường Tín": [],"H. Ứng Hòa": []
  },
  "Đà Nẵng": {
    "Q. Hải Châu": [],"Q. Thanh Khê": [],"Q. Sơn Trà": [],"Q. Ngũ Hành Sơn": [],"Q. Liên Chiểu": [],"Q. Cẩm Lệ": [],"H. Hòa Vang": [],"H. Hoàng Sa": []
  },
  "Hải Phòng": {
    "Q. Hồng Bàng": [],"Q. Lê Chân": [],"Q. Ngô Quyền": [],"Q. Kiến An": [],"Q. Hải An": [],"Q. Đồ Sơn": [],"Q. Dương Kinh": [],"H. An Dương": [],"H. An Lão": [],"H. Bạch Long Vĩ": [],"H. Cát Hải": [],"H. Kiến Thụy": [],"H. Thủy Nguyên": [],"H. Tiên Lãng": [],"H. Vĩnh Bảo": []
  },
  "Cần Thơ": {
    "Q. Ninh Kiều": [],"Q. Ô Môn": [],"Q. Bình Thủy": [],"Q. Cái Răng": [],"Q. Thốt Nốt": [],"H. Phong Điền": [],"H. Cờ Đỏ": [],"H. Thới Lai": [],"H. Vĩnh Thạnh": []
  },
  "Bình Dương": {
    "TP. Thủ Dầu Một": [],"TX. Dĩ An": [],"TX. Thuận An": [],"TX. Bến Cát": [],"TX. Tân Uyên": [],"H. Dầu Tiếng": [],"H. Bàu Bàng": [],"H. Phú Giáo": []
  },
  "Đồng Nai": {
    "TP. Biên Hòa": [],"TP. Long Khánh": [],"H. Nhơn Trạch": [],"H. Long Thành": [],"H. Trảng Bom": [],"H. Thống Nhất": [],"H. Xuân Lộc": [],"H. Cẩm Mỹ": [],"H. Định Quán": [],"H. Tân Phú": [],"H. Vĩnh Cửu": []
  },
  "Bà Rịa - Vũng Tàu": {
    "TP. Vũng Tàu": [],"TP. Bà Rịa": [],"TX. Phú Mỹ": [],"H. Châu Đức": [],"H. Côn Đảo": [],"H. Đất Đỏ": [],"H. Long Điền": [],"H. Xuyên Mộc": []
  },
  "Khánh Hòa": {
    "TP. Nha Trang": [],"TP. Cam Ranh": [],"TX. Ninh Hòa": [],"H. Diên Khánh": [],"H. Khánh Sơn": [],"H. Khánh Vĩnh": [],"H. Cam Lâm": [],"H. Trường Sa": []
  },
  "Lâm Đồng": {
    "TP. Đà Lạt": [],"TP. Bảo Lộc": [],"H. Bảo Lâm": [],"H. Cát Tiên": [],"H. Di Linh": [],"H. Đơn Dương": [],"H. Đức Trọng": [],"H. Lạc Dương": [],"H. Lâm Hà": [],"H. Đạ Huoai": [],"H. Đạ Tẻh": []
  },
  "Đắk Lắk": {
    "TP. Buôn Ma Thuột": [],"TX. Buôn Hồ": [],"H. Ea H'leo": [],"H. Ea Kar": [],"H. Ea Súp": [],"H. Krông Ana": [],"H. Krông Bông": [],"H. Krông Búk": [],"H. Krông Năng": [],"H. Krông Pắc": [],"H. Lắk": [],"H. M'Đrắk": [],"H. Buôn Đôn": [],"H. Cư Kuin": [],"H. Cư M'gar": []
  },
  "Gia Lai": {
    "TP. Pleiku": [],"TX. An Khê": [],"TX. Ayun Pa": [],"H. Chư Păh": [],"H. Chư Prông": [],"H. Chư Sê": [],"H. Đak Đoa": [],"H. Đak Pơ": [],"H. Đức Cơ": [],"H. Ia Grai": [],"H. Ia Pa": [],"H. K'Bang": [],"H. Kông Chro": [],"H. Krông Pa": [],"H. Mang Yang": [],"H. Phú Thiện": [],"H. Chư Pưh": []
  },
  "Hậu Giang": {
    "TP. Vị Thanh": [],"TX. Long Mỹ": [],"TX. Ngã Bảy": [],"H. Châu Thành": [],"H. Châu Thành A": [],"H. Long Mỹ": [],"H. Phụng Hiệp": [],"H. Vị Thủy": []
  },
  "Sóc Trăng": {
    "TP. Sóc Trăng": [],"TX. Ngã Năm": [],"TX. Vĩnh Châu": [],"H. Châu Thành": [],"H. Cù Lao Dung": [],"H. Kế Sách": [],"H. Long Phú": [],"H. Mỹ Tú": [],"H. Mỹ Xuyên": [],"H. Thạnh Trị": [],"H. Trần Đề": []
  },
  "Bạc Liêu": {
    "TP. Bạc Liêu": [],"TX. Giá Rai": [],"H. Đông Hải": [],"H. Hòa Bình": [],"H. Hồng Dân": [],"H. Phước Long": [],"H. Vĩnh Lợi": []
  },
  "Cà Mau": {
    "TP. Cà Mau": [],"H. Cái Nước": [],"H. Đầm Dơi": [],"H. Năm Căn": [],"H. Ngọc Hiển": [],"H. Phú Tân": [],"H. Thới Bình": [],"H. Trần Văn Thời": [],"H. U Minh": []
  },
  "Tây Ninh": {
    "TP. Tây Ninh": [],"TX. Hòa Thành": [],"TX. Trảng Bàng": [],"H. Bến Cầu": [],"H. Châu Thành": [],"H. Dương Minh Châu": [],"H. Gò Dầu": [],"H. Tân Biên": [],"H. Tân Châu": []
  },
  "Bình Phước": {
    "TP. Đồng Xoài": [],"TX. Bình Long": [],"TX. Chơn Thành": [],"TX. Phước Long": [],"H. Bù Đăng": [],"H. Bù Đốp": [],"H. Bù Gia Mập": [],"H. Đồng Phú": [],"H. Hớn Quản": [],"H. Lộc Ninh": [],"H. Phú Riềng": []
  },
  "Ninh Bình": {
    "TP. Ninh Bình": [],"TP. Tam Điệp": [],"H. Gia Viễn": [],"H. Hoa Lư": [],"H. Kim Sơn": [],"H. Nho Quan": [],"H. Yên Khánh": [],"H. Yên Mô": []
  },
  "Nam Định": {
    "TP. Nam Định": [],"H. Giao Thủy": [],"H. Hải Hậu": [],"H. Mỹ Lộc": [],"H. Nam Trực": [],"H. Nghĩa Hưng": [],"H. Trực Ninh": [],"H. Vụ Bản": [],"H. Xuân Trường": [],"H. Ý Yên": []
  },
  "Thái Bình": {
    "TP. Thái Bình": [],"H. Đông Hưng": [],"H. Hưng Hà": [],"H. Kiến Xương": [],"H. Quỳnh Phụ": [],"H. Thái Thụy": [],"H. Tiền Hải": [],"H. Vũ Thư": []
  },
  "Hà Nam": {
    "TP. Phủ Lý": [],"TX. Duy Tiên": [],"H. Bình Lục": [],"H. Kim Bảng": [],"H. Lý Nhân": [],"H. Thanh Liêm": []
  },
  "Hưng Yên": {
    "TP. Hưng Yên": [],"TX. Mỹ Hào": [],"H. Ân Thi": [],"H. Khoái Châu": [],"H. Kim Động": [],"H. Phù Cừ": [],"H. Tiên Lữ": [],"H. Văn Giang": [],"H. Văn Lâm": [],"H. Yên Mỹ": []
  },
  "Hải Dương": {
    "TP. Hải Dương": [],"TX. Chí Linh": [],"TX. Kinh Môn": [],"H. Bình Giang": [],"H. Cẩm Giàng": [],"H. Gia Lộc": [],"H. Kim Thành": [],"H. Nam Sách": [],"H. Ninh Giang": [],"H. Thanh Hà": [],"H. Thanh Miện": [],"H. Tứ Kỳ": []
  },
  "Bắc Ninh": {
    "TP. Bắc Ninh": [],"TX. Từ Sơn": [],"H. Gia Bình": [],"H. Lương Tài": [],"H. Quế Võ": [],"H. Thuận Thành": [],"H. Tiên Du": [],"H. Yên Phong": []
  },
  "Vĩnh Phúc": {
    "TP. Vĩnh Yên": [],"TX. Phúc Yên": [],"H. Bình Xuyên": [],"H. Lập Thạch": [],"H. Sông Lô": [],"H. Tam Đảo": [],"H. Tam Dương": [],"H. Vĩnh Tường": [],"H. Yên Lạc": []
  },
  "Phú Thọ": {
    "TP. Việt Trì": [],"TX. Phú Thọ": [],"H. Cẩm Khê": [],"H. Đoan Hùng": [],"H. Hạ Hòa": [],"H. Lâm Thao": [],"H. Phù Ninh": [],"H. Tam Nông": [],"H. Tân Sơn": [],"H. Thanh Ba": [],"H. Thanh Sơn": [],"H. Thanh Thủy": [],"H. Yên Lập": []
  },
  "Thái Nguyên": {
    "TP. Thái Nguyên": [],"TP. Sông Công": [],"TX. Phổ Yên": [],"H. Đại Từ": [],"H. Định Hóa": [],"H. Đồng Hỷ": [],"H. Phú Bình": [],"H. Phú Lương": [],"H. Võ Nhai": []
  },
  "Bắc Giang": {
    "TP. Bắc Giang": [],"H. Hiệp Hòa": [],"H. Lạng Giang": [],"H. Lục Nam": [],"H. Lục Ngạn": [],"H. Sơn Động": [],"H. Tân Yên": [],"H. Việt Yên": [],"H. Yên Dũng": [],"H. Yên Thế": []
  },
  "Lạng Sơn": {
    "TP. Lạng Sơn": [],"H. Bắc Sơn": [],"H. Bình Gia": [],"H. Cao Lộc": [],"H. Chi Lăng": [],"H. Đình Lập": [],"H. Hữu Lũng": [],"H. Lộc Bình": [],"H. Tràng Định": [],"H. Văn Lãng": [],"H. Văn Quan": []
  },
  "Quảng Ninh": {
    "TP. Hạ Long": [],"TP. Cẩm Phả": [],"TP. Uông Bí": [],"TX. Đông Triều": [],"TX. Quảng Yên": [],"H. Ba Chẽ": [],"H. Bình Liêu": [],"H. Cô Tô": [],"H. Đầm Hà": [],"H. Hải Hà": [],"H. Hoành Bồ": [],"H. Tiên Yên": [],"H. Vân Đồn": []
  },
  "Cao Bằng": {
    "TP. Cao Bằng": [],"H. Bảo Lạc": [],"H. Bảo Lâm": [],"H. Hà Quảng": [],"H. Hạ Lang": [],"H. Hòa An": [],"H. Nguyên Bình": [],"H. Phục Hòa": [],"H. Quảng Hòa": [],"H. Thạch An": [],"H. Thông Nông": [],"H. Trà Lĩnh": [],"H. Trùng Khánh": []
  },
  "Bắc Kạn": {
    "TP. Bắc Kạn": [],"H. Ba Bể": [],"H. Bạch Thông": [],"H. Chợ Đồn": [],"H. Chợ Mới": [],"H. Na Rì": [],"H. Ngân Sơn": [],"H. Pác Nặm": []
  },
  "Hà Giang": {
    "TP. Hà Giang": [],"H. Bắc Mê": [],"H. Bắc Quang": [],"H. Đồng Văn": [],"H. Hoàng Su Phì": [],"H. Mèo Vạc": [],"H. Quản Bạ": [],"H. Quang Bình": [],"H. Vị Xuyên": [],"H. Xín Mần": [],"H. Yên Minh": []
  },
  "Tuyên Quang": {
    "TP. Tuyên Quang": [],"H. Chiêm Hóa": [],"H. Hàm Yên": [],"H. Lâm Bình": [],"H. Na Hang": [],"H. Sơn Dương": [],"H. Yên Sơn": []
  },
  "Yên Bái": {
    "TP. Yên Bái": [],"TX. Nghĩa Lộ": [],"H. Lục Yên": [],"H. Mù Cang Chải": [],"H. Trạm Tấu": [],"H. Trấn Yên": [],"H. Văn Chấn": [],"H. Văn Yên": [],"H. Yên Bình": []
  },
  "Lào Cai": {
    "TP. Lào Cai": [],"TX. Sa Pa": [],"H. Bắc Hà": [],"H. Bảo Thắng": [],"H. Bảo Yên": [],"H. Mường Khương": [],"H. Si Ma Cai": [],"H. Văn Bàn": [],"H. Bát Xát": []
  },
  "Điện Biên": {
    "TP. Điện Biên Phủ": [],"TX. Mường Lay": [],"H. Điện Biên": [],"H. Điện Biên Đông": [],"H. Mường Ảng": [],"H. Mường Chà": [],"H. Mường Nhé": [],"H. Nậm Pồ": [],"H. Tủa Chùa": [],"H. Tuần Giáo": []
  },
  "Lai Châu": {
    "TP. Lai Châu": [],"H. Mường Tè": [],"H. Nậm Nhùn": [],"H. Phong Thổ": [],"H. Sìn Hồ": [],"H. Tam Đường": [],"H. Tân Uyên": [],"H. Than Uyên": []
  },
  "Sơn La": {
    "TP. Sơn La": [],"H. Bắc Yên": [],"H. Mai Sơn": [],"H. Mộc Châu": [],"H. Mường La": [],"H. Phù Yên": [],"H. Quỳnh Nhai": [],"H. Sông Mã": [],"H. Sốp Cộp": [],"H. Thuận Châu": [],"H. Vân Hồ": [],"H. Yên Châu": []
  },
  "Hòa Bình": {
    "TP. Hòa Bình": [],"H. Cao Phong": [],"H. Đà Bắc": [],"H. Kim Bôi": [],"H. Kỳ Sơn": [],"H. Lạc Sơn": [],"H. Lạc Thủy": [],"H. Lương Sơn": [],"H. Mai Châu": [],"H. Tân Lạc": [],"H. Yên Thủy": []
  },
  "Đắk Nông": {
    "TP. Gia Nghĩa": [],"H. Cư Jút": [],"H. Đắk G'Long": [],"H. Đắk Mil": [],"H. Đắk R'Lấp": [],"H. Đắk Song": [],"H. Krông Nô": [],"H. Tuy Đức": []
  }
};
/* eslint-enable */
const STORAGE = {
  reports: "nora_reports_v1",
  dataSource: "nora_data_source_v1",
  usersSyncEndpoint: "nora_users_sync_endpoint_v1",
  usersPendingSync: "nora_users_pending_sync_v1",
  criticalStatePendingSync: "nora_critical_state_pending_sync_v1",
  auth: "nora_auth_v1",
  loginPrefs: "nora_login_prefs_v1",
  users: "nora_users_v1",
  logo: "nora_brand_logo_v1",
  customers: "nora_customers_v1",
  inventoryItems: "nora_inventory_items_v1",
  inventoryTransactions: "nora_inventory_transactions_v1",
  customerFilters: "nora_customer_filters_v1",
  activities: "nora_activities_v1",
  recycleBin: "nora_recycle_bin_v1",
  hrFiles: "nora_hr_files_v1",
  schedule: "nora_schedule_v1",
  customerCareProgress: "nora_customer_care_progress_v1",
  customerCareFilters: "nora_customer_care_filters_v1",
  accountingCashflow: "nora_accounting_cashflow_v1",
  accountingCashflowFilters: "nora_accounting_cashflow_filters_v1",
  accountingAttendance: "nora_accounting_attendance_v1",
  accountingAttendanceSource: "nora_accounting_attendance_source_v1",
  accountingAttendanceFilters: "nora_accounting_attendance_filters_v1",
  accountingServicePayrollFilters: "nora_accounting_service_payroll_filters_v1",
  nurseReportOverrides: "nora_nurse_report_overrides_v1",
  telegramSource: "nora_telegram_source_v1",
  rolePermissions: "nora_role_permissions_v1",
  newsPosts: "nora_news_posts_v1",
  newsPinned: "nora_news_pinned_v1",
  newsEvents: "nora_news_events_v1"
};

const LOGO_CANDIDATES = [
  "/nora-care-logo.png",
  "/nora-care-logo.jpg",
  "/nora-care-logo.jpeg",
  "/nora-care-logo.webp",
  "/logo-nora.png",
  "/logo-nora.jpg",
  "/logo-nora.jpeg",
  "/logo-nora.webp",
  "/logo.png",
  "/logo.jpg",
  "/logo.jpeg",
  "/logo.webp",
  "/nora-logo.png"
];

const today = new Date().toISOString().slice(0, 10);
const ROLES = {
  admin: { label: "Admin", canViewData: true, canViewUsers: true, canSubmitReport: true, canManageUsers: true, canSyncData: true, canExportPdf: true, pageAccess: ["news", "home", "metrics", "hr", "customers", "schedule", "care", "accounting", "inventory", "reports", "workflow", "policy", "activity", "access"] },
  ceo: { label: "CEO", canViewData: true, canViewUsers: true, canSubmitReport: true, canManageUsers: true, canSyncData: true, canExportPdf: true, pageAccess: ["news", "home", "metrics", "hr", "customers", "schedule", "care", "accounting", "inventory", "reports", "workflow", "policy", "activity", "access"] },
  head: { label: "Trưởng bộ phận", canViewData: true, canViewUsers: true, canSubmitReport: true, canManageUsers: false, canSyncData: true, canExportPdf: true, pageAccess: ["news", "home", "metrics", "hr", "customers", "schedule", "care", "accounting", "inventory", "reports", "workflow", "policy", "access"] },
  staff: { label: "Nhân viên", canViewData: true, canViewUsers: false, canSubmitReport: false, canManageUsers: false, canSyncData: false, canExportPdf: false, pageAccess: ["news", "home", "customers", "schedule", "care"] }
};

const APP_PAGE_LABELS = {
  news: "Bảng tin",
  home: "Trang chủ",
  metrics: "Chỉ số",
  hr: "Nhân sự",
  customers: "Telesales",
  schedule: "Lịch khách hàng",
  care: "Chăm sóc khách hàng",
  accounting: "Kế toán",
  inventory: "Quản lí kho",
  reports: "Báo cáo",
  workflow: "Quy trình",
  policy: "Nội quy & cơ chế",
  activity: "Lịch sử hoạt động",
  access: "Nguồn dữ liệu"
};

const APP_PAGE_KEYS = Object.keys(APP_PAGE_LABELS);
const PAGE_SUBDOMAIN_ALIASES = {
  hrm: "hr",
  telesales: "customers",
  kho: "inventory",
  "ke-toan": "accounting",
  "cham-soc": "care"
};

function normalizePageKey(input) {
  const normalized = String(input || "").toLowerCase().trim();
  if (!normalized) return "";
  if (APP_PAGE_KEYS.includes(normalized)) return normalized;
  return PAGE_SUBDOMAIN_ALIASES[normalized] || "";
}

function getPageKeyFromLocation() {
  const hashSegment = window.location.hash.replace(/^#\/?/, "").split(/[/?#]/)[0] || "";
  const pageFromHash = normalizePageKey(hashSegment);
  if (pageFromHash) return pageFromHash;

  const pathSegment = window.location.pathname.replace(/^\/+|\/+$/g, "").split("/")[0] || "";
  const pageFromPath = normalizePageKey(pathSegment);
  if (pageFromPath) return pageFromPath;

  const host = window.location.hostname.toLowerCase();
  const subdomain = host.includes(".") ? host.split(".")[0] : "";
  const pageFromSubdomain = normalizePageKey(subdomain);
  return pageFromSubdomain || "";
}

function getPagePath(pageKey) {
  return pageKey === "home" ? "#/" : `#/${pageKey}`;
}

function syncUrlWithPage(pageKey, options = {}) {
  const { replace = false } = options;
  const desiredHash = getPagePath(pageKey);
  const normalizedCurrentHash = window.location.hash || "#/";
  if (normalizedCurrentHash === desiredHash && window.location.pathname === "/") return;

  const desiredUrl = `/${window.location.search}${desiredHash}`;
  if (replace) {
    window.history.replaceState({ pageKey }, "", desiredUrl);
    return;
  }
  window.history.pushState({ pageKey }, "", desiredUrl);
}

const seedUsers = [
  { id: "u-admin", username: "admin", password: "NORA-ADMIN-2026", fullName: "System Admin", roleKey: "admin", department: "Vận hành", createdAt: Date.now() - 1000 * 60 * 60 * 24 * 30 },
  { id: "u-ceo", username: "ceo", password: "NORA-CEO-2026", fullName: "CEO Demo", roleKey: "ceo", department: "Ban điều hành", createdAt: Date.now() - 1000 * 60 * 60 * 24 * 10 },
  { id: "u-head-tech", username: "head-tech", password: "NORA-HEAD-2026", fullName: "Trưởng BP Kỹ thuật", roleKey: "head", department: "Kỹ thuật", createdAt: Date.now() - 1000 * 60 * 60 * 24 * 8 }
];

const seedReports = [
  { date: today, department: "Kỹ thuật", completion: 88, quality: 91, issues: 2, submitter: "Nguyễn Minh", updatedAt: Date.now() - 3600 * 1000 * 4 },
  { date: today, department: "Kinh doanh", completion: 76, quality: 83, issues: 3, submitter: "Trần Phúc", updatedAt: Date.now() - 3600 * 1000 * 2 },
  { date: today, department: "Nhân sự", completion: 92, quality: 95, issues: 0, submitter: "Lê Chi", updatedAt: Date.now() - 3600 * 1000 }
];

const seedNewsPosts = [
  {
    id: "np-seed-1",
    authorName: "Ban Điều Hành",
    department: "Điều hành",
    createdAt: Date.now() - 1000 * 60 * 60 * 26,
    content: "Ban hành quy trình phối hợp liên phòng ban phiên bản 2.0. Toàn bộ trưởng bộ phận rà soát và phản hồi trước 17:00 hôm nay.",
    tags: ["Quan trọng", "Quy trình"],
    views: 36,
    comments: 12,
    tone: "warn"
  },
  {
    id: "np-seed-2",
    authorName: "Phòng CSKH",
    department: "CSKH",
    createdAt: Date.now() - 1000 * 60 * 60 * 43,
    content: "Đội CSKH đạt điểm hài lòng trung bình 4.8/5 trong tuần. Tiếp tục duy trì phản hồi theo SLA để giữ chất lượng dịch vụ.",
    tags: ["Tích cực", "Hiệu suất"],
    views: 22,
    comments: 6,
    tone: "good"
  }
];

const seedNewsPinned = [
  { id: "pin-1", text: "Hạn chót nộp KPI tháng trước: 17:00 Thứ Sáu.", createdAt: Date.now() - 1000 * 60 * 60 * 12 },
  { id: "pin-2", text: "Đào tạo nội bộ chất lượng dịch vụ: 14:00 Thứ Tư.", createdAt: Date.now() - 1000 * 60 * 60 * 36 },
  { id: "pin-3", text: "Bảo trì đồng bộ dữ liệu: 23:00 - 23:30 tối nay.", createdAt: Date.now() - 1000 * 60 * 60 * 5 }
];

const seedNewsEvents = [
  {
    id: "ne-1",
    date: "2026-04-13",
    time: "09:00",
    title: "Giao ban trưởng bộ phận",
    location: "Phòng họp tầng 3",
    description: "Cập nhật tiến độ tuần và các điểm nghẽn liên phòng ban.",
    votes: { "u-admin": "yes", "u-ceo": "yes", "u-head-tech": "yes" },
    createdAt: Date.now() - 1000 * 60 * 60 * 60
  },
  {
    id: "ne-2",
    date: "2026-04-15",
    time: "14:00",
    title: "Đào tạo kỹ năng xử lý phản hồi",
    location: "Phòng đào tạo",
    description: "Chuẩn hóa cách xử lý phản hồi tiêu cực và escalations.",
    votes: { "u-admin": "yes", "u-head-tech": "no" },
    createdAt: Date.now() - 1000 * 60 * 60 * 30
  },
  {
    id: "ne-3",
    date: "2026-04-18",
    time: "10:30",
    title: "Rà soát tồn kho quý 2",
    location: "Kho trung tâm",
    description: "Kiểm kê mã hàng tồn lâu và kế hoạch xuất luân chuyển.",
    votes: { "u-admin": "yes" },
    createdAt: Date.now() - 1000 * 60 * 60 * 10
  }
];

const seedCustomers = [
  {
    id: "c-001",
    name: "Công ty Minh An",
    contactPerson: "Nguyễn Quốc Minh",
    phone: "0903123456",
    email: "minhan@example.com",
    address: "Quận 1, TP.HCM",
    tier: "VIP",
    status: "Đang cân nhắc",
    owner: "admin",
    source: "Website",
    demand: "Nâng cấp gói KPI và dashboard doanh số",
    note: "Đã gọi lần 1, hẹn gửi proposal trong tuần.",
    updatedAt: Date.now() - 86400000 * 2
  },
  {
    id: "c-002",
    name: "Phòng khám Hồng Phúc",
    contactPerson: "Lê Mỹ Hạnh",
    phone: "0911222333",
    email: "hongphuc@example.com",
    address: "Biên Hòa, Đồng Nai",
    tier: "Tiềm năng",
    status: "Không nghe máy",
    owner: "head-tech",
    source: "Facebook",
    demand: "Tích hợp báo cáo chất lượng dịch vụ",
    note: "Đã gọi 2 lần chưa bắt máy, cần chăm lại vào thứ 2.",
    updatedAt: Date.now() - 86400000
  },
  {
    id: "c-003",
    name: "Nhà thuốc Đức Tín",
    contactPerson: "Trần Gia Khánh",
    phone: "0988777666",
    email: "ductin@example.com",
    address: "Cần Thơ",
    tier: "Tiêu chuẩn",
    status: "Đã ký",
    owner: "ceo",
    source: "Giới thiệu",
    demand: "Theo dõi hiệu suất bán hàng theo khu vực",
    note: "Đã ký hợp đồng 12 tháng, đang onboarding.",
    updatedAt: Date.now() - 3600000 * 6
  }
];

const seedInventoryItems = [
  {
    id: "inv-001",
    productCode: "VT-001",
    productName: "Găng tay y tế",
    purchasePrice: 12000,
    salePrice: 17000,
    quantity: 480,
    alertThreshold: 200,
    status: "active",
    updatedAt: Date.now() - 1000 * 60 * 60 * 24
  },
  {
    id: "inv-014",
    productCode: "VT-014",
    productName: "Khẩu trang 4 lớp",
    purchasePrice: 1500,
    salePrice: 2500,
    quantity: 120,
    alertThreshold: 150,
    status: "active",
    updatedAt: Date.now() - 1000 * 60 * 60 * 5
  },
  {
    id: "inv-022",
    productCode: "VT-022",
    productName: "Dung dịch sát khuẩn 500ml",
    purchasePrice: 28000,
    salePrice: 39000,
    quantity: 95,
    alertThreshold: 90,
    status: "active",
    updatedAt: Date.now() - 1000 * 60 * 60 * 2
  }
];

const seedAccountingCashflow = [
  {
    id: "cf-001",
    date: today,
    type: "income",
    voucherCode: `PT-${today.replaceAll("-", "").slice(2)}-01`,
    category: "Thu từ khách hàng",
    counterparty: "Phan Bảo Trâm",
    content: "Thanh toán gói chăm sóc mẹ bầu 10 buổi",
    amount: 12450000,
    method: "Chuyển khoản",
    creator: "Ngọc Anh",
    status: "approved",
    createdAt: Date.now() - 1000 * 60 * 60 * 6
  },
  {
    id: "cf-002",
    date: today,
    type: "expense",
    voucherCode: `PC-${today.replaceAll("-", "").slice(2)}-02`,
    category: "Hoàn ứng điều dưỡng",
    counterparty: "Yến",
    content: "Hoàn ứng chi phí di chuyển ca tối",
    amount: 1800000,
    method: "Tiền mặt",
    creator: "Hà My",
    status: "pending",
    createdAt: Date.now() - 1000 * 60 * 60 * 4
  },
  {
    id: "cf-003",
    date: `${today.slice(0, 8)}12`,
    type: "expense",
    voucherCode: `PC-${today.slice(2, 4)}${today.slice(5, 7)}12-04`,
    category: "Chi vật tư",
    counterparty: "Nhà cung cấp MedCare",
    content: "Thanh toán vật tư chăm sóc mẹ và bé đợt 2",
    amount: 9650000,
    method: "Chuyển khoản",
    creator: "Thu Trang",
    status: "approved",
    createdAt: Date.now() - 1000 * 60 * 60 * 24
  }
];

const seedAccountingAttendance = [
  {
    id: "at-001",
    employeeCode: "NR001",
    employeeName: "Yến",
    department: "Điều dưỡng",
    date: today,
    checkin: "08:00",
    checkout: "17:15",
    workHours: 8.5,
    overtimeHours: 0.5,
    lateMinutes: 0
  },
  {
    id: "at-002",
    employeeCode: "NR002",
    employeeName: "Quyền",
    department: "Tư vấn",
    date: today,
    checkin: "08:12",
    checkout: "17:05",
    workHours: 8,
    overtimeHours: 0,
    lateMinutes: 12
  },
  {
    id: "at-003",
    employeeCode: "NR003",
    employeeName: "Hồ Trang",
    department: "Telesales",
    date: today,
    checkin: "07:55",
    checkout: "18:10",
    workHours: 9,
    overtimeHours: 1,
    lateMinutes: 0
  }
];

const app = document.querySelector("#app");

const seedSchedule = [
  { id: "sc-001", registrationDate: "2026-04-01", appointmentTime: "10h", customerName: "Trần Thuý Anh", phone: "328897997", address: "Số nhà 52 L11 khu đô thị Louis Tân Mai Hoàng Mai", motherAge: "", birthHistory: "", babyBirthday: "", service: "Chăm sóc mẹ bầu", stage: "Mẹ bầu 20w", motherCondition: "Bầu 20 tuần, đau mỏi nhiều phần lưng, mới trải nghiệm gói chăm sóc bầu 249k/ 90 phút", priority: "", babyCondition: "", consultant: "Quyền", nurse: "Yến", experiencePrice: 249000, sessionDuration: "90p", saleStaff: "Hồ Trang", source: "Facebook", contractAmount: 249000, status: "confirmed", note: "", createdAt: Date.now() - 86400000 * 11, updatedAt: Date.now() - 86400000 * 11 },
  { id: "sc-002", registrationDate: "2026-04-01", appointmentTime: "10h", customerName: "Phạm Thanh Bình", phone: "912265555", address: "Chung cư brg diamond residence 25 lê văn lương - thanh xuân - hn", motherAge: "", birthHistory: "", babyBirthday: "", service: "Chăm sóc mẹ & bé", stage: "Em bé", motherCondition: "Tắm bé 45p, bé được hơn 10 ngày tuổi, khách quan tâm gói tắm chăm sóc bé 1 tháng, hỏi thêm về các vấn đề cảm dịch âm", priority: "10 ngày tuổi", babyCondition: "", consultant: "Phương", nurse: "", experiencePrice: 99000, sessionDuration: "45p", saleStaff: "Hồ Trang", source: "Zalo", contractAmount: 99000, status: "confirmed", note: "", createdAt: Date.now() - 86400000 * 11, updatedAt: Date.now() - 86400000 * 11 },
  { id: "sc-003", registrationDate: "2026-03-31", appointmentTime: "10h30 - 10h40", customerName: "Ma Kết", phone: "988211188", address: "Số 36 Hoàng Cầu, Ô Chợ Dừa", motherAge: "", birthHistory: "", babyBirthday: "", service: "Chăm sóc mẹ bầu", stage: "Mẹ bầu 19w", motherCondition: "Bầu 19 tuần, khách muốn massage chăm sóc body bầu, k massa mặt, tư vấn trải nghiệm 90 phút 249k", priority: "", babyCondition: "", consultant: "Thanh", nurse: "Linh", experiencePrice: 249000, sessionDuration: "90p", saleStaff: "", source: "Giới thiệu", contractAmount: 0, status: "cancelled", note: "", createdAt: Date.now() - 86400000 * 12, updatedAt: Date.now() - 86400000 * 12 },
  { id: "sc-004", registrationDate: "2026-03-25", appointmentTime: "14h", customerName: "Hà Trương", phone: "989129899", address: "Số 6 ngõ 97 Khương Trung, Thanh Xuân, Hà Nội", motherAge: "", birthHistory: "", babyBirthday: "", service: "Chăm sóc mẹ bầu", stage: "Mẹ bầu 23w", motherCondition: "Bầu 23 tuần, trải nghiệm 1 buổi chăm sóc mẹ bầu 120p 349k, đặt lịch lại 2 tuần sau", priority: "", babyCondition: "", consultant: "Thanh", nurse: "Nhung", experiencePrice: 349000, sessionDuration: "120p", saleStaff: "", source: "Zalo", contractAmount: 0, status: "cancelled", note: "", createdAt: Date.now() - 86400000 * 18, updatedAt: Date.now() - 86400000 * 18 },
  { id: "sc-005", registrationDate: "2026-04-02", appointmentTime: "15h30", customerName: "Dotty Nguyen", phone: "988080861", address: "Tòa hh1b chung cư meco 102 trường chính", motherAge: "", birthHistory: "", babyBirthday: "", service: "Chăm sóc mẹ bầu", stage: "Mẹ bầu 17w", motherCondition: "Bầu 17 tuần, khu vực Trường Chính, hay bị nhức mỏi, cơ rút thêm phần chân, trải nghiệm buổi 249k/ 90p", priority: "", babyCondition: "", consultant: "Phương", nurse: "Linh", experiencePrice: 249000, sessionDuration: "90p", saleStaff: "Hồ Trang", source: "Facebook", contractAmount: 16749000, status: "confirmed", note: "", createdAt: Date.now() - 86400000 * 10, updatedAt: Date.now() - 86400000 * 10 },
  { id: "sc-006", registrationDate: "2026-04-02", appointmentTime: "16h30", customerName: "c Dương", phone: "0965899999", address: "Phòng 1902, Grandeur Place - Sảnh A, 138B Giảng Võ, TP Hà Nội", motherAge: "", birthHistory: "", babyBirthday: "", service: "Chăm sóc mẹ bầu", stage: "Mẹ bầu 15w", motherCondition: "Bầu 15 tuần, đang bị bó cơ, đau đầu. Thời gian: hơn 4h đến 4h30", priority: "", babyCondition: "", consultant: "Quyền", nurse: "Khuyên", experiencePrice: 249000, sessionDuration: "90p", saleStaff: "Hồ Trang", source: "Facebook", contractAmount: 39349000, status: "confirmed", note: "", createdAt: Date.now() - 86400000 * 10, updatedAt: Date.now() - 86400000 * 10 },
  { id: "sc-007", registrationDate: "2026-04-02", appointmentTime: "10h", customerName: "Lê Thu Thủy", phone: "912351988", address: "P709 25T2 Nguyễn Thị Thập, Yên Hòa, Hà Nội", motherAge: "38 tuổi", birthHistory: "", babyBirthday: "", service: "Chăm sóc mẹ bầu", stage: "Mẹ bầu 37w", motherCondition: "Bầu 37 tuần, đợt này bị mỏi ng, khó ngủ, quan tâm gói chăm sóc mẹ và bé sau sinh, mới trải nghiệm chăm sóc bầu 249k/90p", priority: "", babyCondition: "", consultant: "Quyền", nurse: "Linh", experiencePrice: 249000, sessionDuration: "90p", saleStaff: "Hồ Trang", source: "Zalo", contractAmount: 5249000, status: "completed", note: "", createdAt: Date.now() - 86400000 * 10, updatedAt: Date.now() - 86400000 * 10 },
  { id: "sc-008", registrationDate: "2026-04-03", appointmentTime: "10h", customerName: "Mai Dương", phone: "989606356", address: "188 Đường Bưởi", motherAge: "", birthHistory: "", babyBirthday: "", service: "Chăm sóc mẹ bầu", stage: "Mẹ bầu 37w", motherCondition: "Khách ở Hoàng Quốc Việt, bầu 37 tuần, đau mỏi nhiều", priority: "", babyCondition: "", consultant: "Thanh", nurse: "Phương", experiencePrice: 249000, sessionDuration: "90p", saleStaff: "Hồ Trang", source: "Giới thiệu", contractAmount: 249000, status: "pending", note: "", createdAt: Date.now() - 86400000 * 9, updatedAt: Date.now() - 86400000 * 9 },
  { id: "sc-009", registrationDate: "2026-04-02", appointmentTime: "1h30", customerName: "Nguyễn Thúy Nga", phone: "366319268", address: "Số 49 ngách 1/66 đường văn bội, Bắc Từ Liêm Hà Nội", motherAge: "", birthHistory: "", babyBirthday: "", service: "Chăm sóc mẹ bầu", stage: "Mẹ bầu 33w", motherCondition: "Gần học viện tài chính, bác sĩ liêm hà nội, bầu 33 tuần, phù chân mỏi lưng, tư vấn buổi trải nghiệm 249k/ 90 phút", priority: "", babyCondition: "", consultant: "Thanh", nurse: "Nhung", experiencePrice: 249000, sessionDuration: "90p", saleStaff: "Hồ Trang", source: "Facebook", contractAmount: 2749000, status: "completed", note: "", createdAt: Date.now() - 86400000 * 10, updatedAt: Date.now() - 86400000 * 10 },
  { id: "sc-010", registrationDate: "2026-04-05", appointmentTime: "09h30", customerName: "Phan Bảo Trâm", phone: "0911288668", address: "KĐT Vinhomes Smart City, Nam Từ Liêm, Hà Nội", motherAge: "31 tuổi", birthHistory: "Con so", babyBirthday: "", service: "Chăm sóc mẹ bầu", stage: "Mẹ bầu 28w", motherCondition: "Mất ngủ 2 tuần gần đây, đau thắt lưng dưới, quan tâm liệu trình 5 buổi", priority: "", babyCondition: "", consultant: "Quyền", nurse: "Yến", experiencePrice: 249000, sessionDuration: "90p", saleStaff: "Hồ Trang", source: "Google Ads", contractAmount: 12450000, status: "confirmed", note: "Đã cọc 1 triệu", createdAt: Date.now() - 86400000 * 7, updatedAt: Date.now() - 86400000 * 7 },
  { id: "sc-011", registrationDate: "2026-04-06", appointmentTime: "11h00", customerName: "Đỗ Minh Châu", phone: "0983388123", address: "Ngõ 67 Phùng Khoang, Thanh Xuân, Hà Nội", motherAge: "29 tuổi", birthHistory: "Con 1", babyBirthday: "", service: "Tư vấn chăm sóc sau sinh", stage: "Mẹ bầu 35w", motherCondition: "Khách muốn tìm hiểu combo mẹ và bé sau sinh tại nhà", priority: "", babyCondition: "", consultant: "Thanh", nurse: "Nhung", experiencePrice: 149000, sessionDuration: "60p", saleStaff: "Linh Phạm", source: "Tiktok", contractAmount: 0, status: "pending", note: "Hẹn gọi lại 20h tối", createdAt: Date.now() - 86400000 * 6, updatedAt: Date.now() - 86400000 * 6 },
  { id: "sc-012", registrationDate: "2026-04-07", appointmentTime: "14h30", customerName: "Ngô Quỳnh Mai", phone: "0902333109", address: "Ecohome 3, Đông Ngạc, Bắc Từ Liêm, Hà Nội", motherAge: "34 tuổi", birthHistory: "Con 2", babyBirthday: "15 ngày tuổi", service: "Tắm bé + chăm sóc mẹ sau sinh", stage: "Em bé", motherCondition: "Sau sinh mổ, cần hỗ trợ chăm sóc mẹ và bé trong 1 tháng", priority: "15 ngày tuổi", babyCondition: "Bé ngủ không sâu giấc", consultant: "Phương", nurse: "Khuyên", experiencePrice: 199000, sessionDuration: "75p", saleStaff: "Hồ Trang", source: "Giới thiệu", contractAmount: 8900000, status: "completed", note: "Đã ký gói 1 tháng", createdAt: Date.now() - 86400000 * 5, updatedAt: Date.now() - 86400000 * 4 },
  { id: "sc-013", registrationDate: "2026-04-08", appointmentTime: "16h00", customerName: "Vũ Thị Hà", phone: "0977112288", address: "CT3A Mễ Trì Thượng, Nam Từ Liêm, Hà Nội", motherAge: "27 tuổi", birthHistory: "Con so", babyBirthday: "", service: "Massage bầu giảm đau mỏi", stage: "Mẹ bầu 24w", motherCondition: "Đau vai gáy kéo dài, bắp chân hay bị chuột rút khi ngủ", priority: "", babyCondition: "", consultant: "Quyền", nurse: "Linh", experiencePrice: 249000, sessionDuration: "90p", saleStaff: "Ngọc Anh", source: "Facebook", contractAmount: 0, status: "cancelled", note: "Khách bận công tác, xin dời lịch", createdAt: Date.now() - 86400000 * 4, updatedAt: Date.now() - 86400000 * 3 },
  { id: "sc-014", registrationDate: "2026-04-09", appointmentTime: "18h30", customerName: "Lương Kim Oanh", phone: "0935661200", address: "An Bình City, Cổ Nhuế 1, Bắc Từ Liêm, Hà Nội", motherAge: "32 tuổi", birthHistory: "Con 1", babyBirthday: "", service: "Combo chăm sóc mẹ bầu chuyên sâu", stage: "Mẹ bầu 30w", motherCondition: "Phù nhẹ bàn chân, đau cột sống thắt lưng, mong muốn chăm sóc định kỳ đến lúc sinh", priority: "", babyCondition: "", consultant: "Thanh", nurse: "Yến", experiencePrice: 349000, sessionDuration: "120p", saleStaff: "Hồ Trang", source: "Website", contractAmount: 15900000, status: "confirmed", note: "Đã gửi hợp đồng điện tử", createdAt: Date.now() - 86400000 * 3, updatedAt: Date.now() - 86400000 * 2 }
];
app.innerHTML = `
  <div class="app" id="dashboardRoot">
    <header class="topbar card">
      <div class="header-actions post-login-only hidden">
        <button class="menu-toggle" id="menuToggle" type="button" aria-label="Mở menu" title="Mở menu">☰</button>
        <button class="back-toggle" id="backBtn" type="button" aria-label="Quay lại trang trước" title="Quay lại">←</button>
      </div>
      <div class="brand-block" id="brandBlock">
        <input id="logoUpload" class="hidden" type="file" accept="image/*" />
        <div class="brand-content">
          <img class="brand-title-logo hidden" id="brandTitleLogo" alt="Nora Care" />
          <h1 id="brandTitleText">Nora Care</h1>
          <p>Chào mừng bạn tới trang quản trị hệ thống của Nora Care</p>
        </div>
      </div>
    </header>

    <div class="menu-overlay post-login-only hidden" id="menuOverlay"></div>
    <nav class="menu-drawer post-login-only hidden" id="appMenu" aria-label="Menu chức năng">
      <div class="menu-brand">
        <span class="menu-brand-icon">🌿</span>
        <div>
          <div class="menu-brand-name">Nora Care</div>
          <div class="menu-brand-sub">Hệ thống quản trị</div>
        </div>
      </div>
      <div class="menu-section-label">TỔNG QUAN</div>
      <button class="menu-item" type="button" data-page="news"><span class="mi-icon">📰</span><span>Bảng tin</span></button>
      <button class="menu-item active" type="button" data-page="home"><span class="mi-icon">🏠</span><span>Trang chủ</span></button>
      <button class="menu-item" type="button" data-page="metrics"><span class="mi-icon">📊</span><span>Chỉ số</span></button>
      <div class="menu-section-label">QUẢN LÍ</div>
      <button class="menu-item" type="button" data-page="hr"><span class="mi-icon">👥</span><span>Nhân sự</span></button>
      <button class="menu-item" type="button" data-page="customers"><span class="mi-icon">🤝</span><span>Telesales</span></button>
      <button class="menu-item" type="button" data-page="schedule"><span class="mi-icon">📅</span><span>Lịch khách hàng</span></button>
      <button class="menu-item" type="button" data-page="care"><span class="mi-icon">💆</span><span>Chăm sóc khách hàng</span></button>
      <button class="menu-item" type="button" data-page="accounting"><span class="mi-icon">🧾</span><span>Kế toán</span></button>
      <button class="menu-item" type="button" data-page="inventory"><span class="mi-icon">📦</span><span>Quản lí kho</span></button>
      <div class="menu-section-label">HỆ THỐNG</div>
      <button class="menu-item" type="button" data-page="reports"><span class="mi-icon">🗂️</span><span>Báo cáo</span></button>
      <button class="menu-item" type="button" data-page="workflow"><span class="mi-icon">🔄</span><span>Quy trình</span></button>
      <button class="menu-item" type="button" data-page="policy"><span class="mi-icon">📋</span><span>Nội quy & cơ chế</span></button>
      <button class="menu-item" type="button" data-page="activity"><span class="mi-icon">🕐</span><span>Lịch sử hoạt động</span></button>
      <button class="menu-item" type="button" data-page="access"><span class="mi-icon">🔌</span><span>Nguồn dữ liệu</span></button>
      <button class="menu-logout-btn" id="menuLogoutBtn" type="button"><span class="mi-icon">🚪</span><span>Đăng xuất</span></button>
    </nav>

    <div class="layout">
      <aside class="card side auth-panel">
        <section id="authSection">
          <h3>Đăng nhập</h3>
          <div class="form-grid">
            <div>
              <label>Tên đăng nhập</label>
              <input id="loginUsername" placeholder="Nhập username" autocomplete="username" />
            </div>
            <div>
              <label>Mật khẩu</label>
              <div class="password-field-wrap">
                <input id="loginPassword" type="password" placeholder="Nhập mật khẩu" autocomplete="current-password" />
                <button class="password-toggle-btn" id="toggleLoginPasswordBtn" type="button" aria-label="Hiện mật khẩu" title="Hiện/ẩn mật khẩu">👁</button>
              </div>
            </div>
            <label class="remember-login-row">
              <input id="loginRemember" type="checkbox" />
              <span>Nhớ mật khẩu</span>
            </label>
            <div class="login-actions">
              <button class="btn" id="loginBtn">Đăng nhập</button>
              <button class="btn warn hidden" id="logoutBtn">Đăng xuất</button>
            </div>
            <div class="alert" id="authMessage"></div>
          </div>
        </section>
      </aside>

      <main class="main post-login-only hidden" id="mainContent">
        <section class="card section app-page hidden" data-page="news" id="newsSection">
          <div class="news-header">
            <div>
              <h3>Bảng tin nội bộ</h3>
              <p class="muted">Không gian thông báo chung và tin tức vận hành theo phong cách mạng xã hội.</p>
            </div>
            <span class="news-date">Cập nhật: ${new Date().toLocaleDateString("vi-VN")}</span>
          </div>

          <div class="news-layout">
            <div class="news-feed-col">
              <article class="news-composer card">
                <div class="news-composer-top">
                  <div class="news-avatar">NC</div>
                  <button class="news-compose-input" type="button" id="newsComposeInputBtn" data-news-action="open-composer">Chia sẻ thông báo với toàn hệ thống...</button>
                </div>
                <div class="news-composer-editor hidden" id="newsComposerEditor">
                  <textarea id="newsComposerText" rows="3" placeholder="Nhập nội dung thông báo..."></textarea>
                  <div class="news-attach-uploader">
                    <label>Đính kèm ảnh/file</label>
                    <input id="newsComposerAttachmentInput" type="file" multiple />
                    <div id="newsComposerAttachmentList" class="news-editor-attachment-list"></div>
                  </div>
                  <div class="news-composer-tools">
                    <div>
                      <label>Phòng ban</label>
                      <select id="newsComposerDept">
                        <option value="Toàn hệ thống">Toàn hệ thống</option>
                        ${DEPARTMENTS.map((dept) => `<option value="${dept}">${dept}</option>`).join("")}
                      </select>
                    </div>
                    <div>
                      <label>Ngày sự kiện</label>
                      <input id="newsComposerEventDate" type="date" value="${today}" />
                    </div>
                    <label class="news-important-toggle">
                      <input id="newsComposerImportant" type="checkbox" />
                      <span>Đánh dấu thông báo ghim</span>
                    </label>
                  </div>
                </div>
                <div class="news-composer-actions">
                  <button class="btn secondary" id="newsSubmitPostBtn" type="button" data-news-action="submit-post">Đăng thông báo</button>
                  <button class="btn secondary" type="button" data-news-action="create-event">Tạo sự kiện</button>
                  <button class="btn secondary" type="button" data-news-action="attach-department">Gắn phòng ban</button>
                </div>
              </article>

              <div class="news-story-row">
                <article class="news-story card"><span>Marketing</span><strong>Chiến dịch Q2</strong></article>
                <article class="news-story card"><span>Kinh doanh</span><strong>Tăng trưởng 12%</strong></article>
                <article class="news-story card"><span>Điều hành</span><strong>Lịch cao điểm</strong></article>
                <article class="news-story card"><span>CSKH</span><strong>Điểm 4.8/5</strong></article>
              </div>

              <div id="newsEventFeedList"></div>
              <div id="newsPostList"></div>
            </div>

            <aside class="news-side-col">
              <article class="card news-side-card">
                <h4>Thông báo ghim</h4>
                <ul id="newsPinnedList"></ul>
              </article>

              <article class="card news-side-card">
                <h4>Sự kiện sắp tới</h4>
                <div id="newsEventsList"></div>
              </article>

              <article class="card news-side-card">
                <h4>Xu hướng nội bộ</h4>
                <div id="newsTrendList"></div>
              </article>
            </aside>
          </div>

          <div id="newsEventModal" class="customer-modal hidden">
            <div class="customer-modal-backdrop" id="newsEventModalBackdrop"></div>
            <div class="customer-modal-panel news-event-modal-panel">
              <h3>Quản lý sự kiện và xác nhận tham gia</h3>
              <p class="muted" style="margin-top:4px;">Tạo sự kiện, hẹn lịch và thu vote xác nhận tham gia từ nhân sự.</p>

              <div class="form-grid">
                <div>
                  <label>Tiêu đề sự kiện</label>
                  <input id="newsEventTitle" placeholder="Ví dụ: Họp điều phối tuần" />
                </div>
                <div>
                  <label>Ngày hẹn</label>
                  <input id="newsEventDate" type="date" value="${today}" />
                </div>
                <div>
                  <label>Giờ hẹn</label>
                  <input id="newsEventTime" type="time" value="09:00" />
                </div>
                <div>
                  <label>Địa điểm</label>
                  <input id="newsEventLocation" placeholder="Phòng họp / Online" />
                </div>
                <div style="grid-column:1 / -1;">
                  <label>Nội dung lịch hẹn</label>
                  <textarea id="newsEventDescription" rows="2" placeholder="Mục tiêu, thành phần, ghi chú..."></textarea>
                </div>
                <div style="grid-column:1 / -1;" class="news-attach-uploader">
                  <label>Đính kèm ảnh/file sự kiện</label>
                  <input id="newsEventAttachmentInput" type="file" multiple />
                  <div id="newsEventAttachmentList" class="news-editor-attachment-list"></div>
                </div>
                <div class="login-actions" style="grid-column:1 / -1;">
                  <button class="btn secondary" id="saveNewsEventBtn" type="button">Lưu sự kiện</button>
                  <button class="btn warn" id="closeNewsEventModalBtn" type="button">Đóng</button>
                </div>
                <div class="alert" id="newsEventModalStatus" style="grid-column:1 / -1;"></div>
              </div>
            </div>
          </div>

          <div id="newsVoteDetailModal" class="customer-modal hidden">
            <div class="customer-modal-backdrop" id="newsVoteDetailModalBackdrop"></div>
            <div class="customer-modal-panel news-vote-detail-panel">
              <h3 id="newsVoteDetailTitle">Chi tiết bình chọn</h3>
              <div id="newsVoteDetailList" class="news-vote-detail-list"></div>
              <div class="login-actions" style="margin-top:10px;">
                <button class="btn warn" id="closeNewsVoteDetailBtn" type="button">Đóng</button>
              </div>
            </div>
          </div>

          <div id="newsContentDetailModal" class="customer-modal hidden">
            <div class="customer-modal-backdrop" id="newsContentDetailBackdrop"></div>
            <div class="customer-modal-panel news-content-detail-panel">
              <h3 id="newsContentDetailTitle">Chi tiết bài đăng</h3>
              <p class="muted" id="newsContentDetailMeta"></p>
              <p id="newsContentDetailText"></p>
              <div id="newsContentDetailImagePreview" class="news-detail-image-preview hidden"></div>
              <div id="newsContentDetailAttachments" class="news-detail-attachments"></div>
              <div class="login-actions" style="margin-top:10px;">
                <button class="btn warn" id="closeNewsContentDetailBtn" type="button">Đóng</button>
              </div>
            </div>
          </div>
        </section>

        <section class="card section app-page" data-page="home">
          <h3>Trang chủ điều hành</h3>
          <p class="muted">Theo dõi doanh số, tiến độ KPI và chỉ số tổng quan các bộ phận theo thời gian.</p>
          <div class="time-filter form-grid">
            <div>
              <label>Kiểu lọc thời gian</label>
              <select id="timePreset">
                <option value="today">Hôm nay</option>
                <option value="7d">7 ngày gần nhất</option>
                <option value="30d">30 ngày gần nhất</option>
                <option value="thisMonth">Tháng hiện tại</option>
                <option value="custom">Tùy chọn khoảng ngày</option>
              </select>
            </div>
            <div>
              <label>Từ ngày</label>
              <input id="filterStart" type="date" />
            </div>
            <div>
              <label>Đến ngày</label>
              <input id="filterEnd" type="date" />
            </div>
            <button class="btn secondary" id="applyFilterBtn" type="button">Áp dụng bộ lọc</button>
            <button class="btn secondary" id="resetFilterBtn" type="button">Đặt lại hôm nay</button>
            <div class="muted" id="filterSummary">Bộ lọc hiện tại: Hôm nay</div>
          </div>
          <div class="kpis">
            <article class="kpi"><div class="label">Báo cáo trong kỳ</div><div class="value" id="kpiReports">0</div></article>
            <article class="kpi"><div class="label">Hoàn thành TB</div><div class="value" id="kpiCompletion">0%</div></article>
            <article class="kpi"><div class="label">Chất lượng TB</div><div class="value" id="kpiQuality">0</div></article>
            <article class="kpi"><div class="label">Tổng sự cố</div><div class="value" id="kpiIssues">0</div></article>
          </div>
          <div class="revenue-overview">
            <article class="overview-card revenue">
              <div class="label">Doanh thu ngày</div>
              <div class="value" id="homeDailyRevenue">0 đ</div>
              <div class="muted" id="homeDailyRevenueMeta">--</div>
            </article>
            <article class="overview-card revenue">
              <div class="label">Doanh thu tháng</div>
              <div class="value" id="homeMonthlyRevenue">0 đ</div>
              <div class="muted" id="homeMonthlyRevenueMeta">--</div>
            </article>
            <article class="overview-card revenue">
              <div class="label">Doanh thu tổng</div>
              <div class="value" id="homeTotalRevenue">0 đ</div>
              <div class="muted" id="homeTotalRevenueMeta">--</div>
            </article>
          </div>
          <section class="card section" style="margin-top:12px;">
            <h3 style="margin-bottom:10px;">Biểu đồ doanh số và tiến độ KPI</h3>
            <canvas id="homeRevenueChart" height="110"></canvas>
          </section>
          <div class="tables" style="margin-top:12px;">
            <table>
              <thead>
                <tr><th>Bộ phận</th><th>Chỉ số tổng hợp</th><th>Doanh số ước tính trong kỳ</th><th>Mức cảnh báo</th></tr>
              </thead>
              <tbody id="homeDeptBody"></tbody>
            </table>
          </div>
          <div class="home-links">
            <button class="btn secondary" type="button" data-go-page="metrics">Xem chỉ số chi tiết</button>
            <button class="btn secondary" type="button" data-go-page="workflow">Đi đến quy trình làm việc</button>
            <button class="btn secondary" type="button" data-go-page="access">Quản trị quyền truy cập</button>
          </div>
        </section>

        <section class="card section app-page hidden" data-page="metrics">
          <h3>Chỉ số</h3>
          <p class="muted" style="margin:4px 0 10px;">Tổng hợp hiệu suất theo phòng ban: Marketing, Telesale, Tư vấn và Điều dưỡng.</p>
          <div class="metrics-filter form-grid" style="margin-top:8px;margin-bottom:10px;">
            <div>
              <label>Từ ngày</label>
              <input id="metricsStartDate" type="date" />
            </div>
            <div>
              <label>Đến ngày</label>
              <input id="metricsEndDate" type="date" />
            </div>
            <div>
              <label>Phòng ban</label>
              <select id="metricsDepartmentFilter">
                <option value="all">Tất cả phòng ban</option>
                <option value="marketing">Marketing</option>
                <option value="telesale">Telesale</option>
                <option value="consultant">Tư vấn</option>
                <option value="nurse">Điều dưỡng</option>
              </select>
            </div>
            <button class="btn secondary" id="applyMetricsFilterBtn" type="button">Áp dụng</button>
            <button class="btn secondary" id="resetMetricsFilterBtn" type="button">Đặt lại</button>
            <div class="muted" id="metricsFilterSummary" style="display:flex;align-items:center;font-size:0.83rem;">Bộ lọc: tất cả thời gian, tất cả phòng ban</div>
          </div>
          <div class="metrics-dept-grid" id="metricsDeptGrid">
            <article class="metrics-dept-card">
              <h4>Marketing</h4>
              <div class="metrics-dept-row"><span>Lượng tương tác</span><strong id="mkInteractions">0</strong></div>
              <div class="metrics-dept-row"><span>SĐT từ Marketing</span><strong id="mkPhones">0</strong></div>
              <div class="metrics-dept-row"><span>Tỉ lệ lead digital</span><strong id="mkLeadRate">0%</strong></div>
            </article>
            <article class="metrics-dept-card">
              <h4>Telesale</h4>
              <div class="metrics-dept-row"><span>Lượng lịch trải nghiệm</span><strong id="tsAppointments">0</strong></div>
              <div class="metrics-dept-row"><span>Tỉ lệ đặt lịch</span><strong id="tsBookingRate">0%</strong></div>
              <div class="metrics-dept-row"><span>Tỉ lệ hủy lịch</span><strong id="tsCancelRate">0%</strong></div>
            </article>
            <article class="metrics-dept-card">
              <h4>Tư vấn</h4>
              <div class="metrics-dept-row"><span>Doanh số</span><strong id="tvRevenue">0 đ</strong></div>
              <div class="metrics-dept-row"><span>Tỉ lệ ký</span><strong id="tvSignRate">0%</strong></div>
              <div class="metrics-dept-row"><span>HĐ trung bình</span><strong id="tvAvgContract">0 đ</strong></div>
            </article>
            <article class="metrics-dept-card">
              <h4>Điều dưỡng</h4>
              <div class="metrics-dept-row"><span>KH đang chăm sóc</span><strong id="ddCaring">0</strong></div>
              <div class="metrics-dept-row"><span>Điểm dịch vụ</span><strong id="ddServiceScore">0/5</strong></div>
              <div class="metrics-dept-row"><span>Tỉ lệ hoàn tất</span><strong id="ddCompleteRate">0%</strong></div>
            </article>
          </div>
          <section class="card section" style="margin-top:10px;">
            <h4 style="margin:0 0 8px;">Biểu đồ KPI phòng ban (đồng bộ theo bộ lọc)</h4>
            <canvas id="metricsDeptKpiChart" height="110"></canvas>
          </section>
          <section class="card section" style="margin-top:10px;">
            <h4 style="margin:0 0 8px;">Tải &amp; đồng bộ dữ liệu chỉ số</h4>
            <div class="metrics-sync-grid">
              <div>
                <label>File dữ liệu (.json / .csv)</label>
                <input id="metricsDataFile" type="file" accept=".json,.csv" />
              </div>
              <button class="btn secondary" id="metricsImportFileBtn" type="button">Tải file vào dữ liệu</button>
              <div>
                <label>Google Sheets / CSV URL</label>
                <input id="metricsSheetUrl" placeholder="https://docs.google.com/... hoặc link csv/json" />
              </div>
              <button class="btn secondary" id="metricsSyncSheetBtn" type="button">Đồng bộ từ link</button>
            </div>
            <p class="muted" style="margin-top:6px;font-size:0.8rem;">Ưu tiên JSON array. Với Google Sheets, nên dùng link publish CSV để đồng bộ ổn định.</p>
          </section>
          <div class="tables" style="margin-top:12px;">
            <table>
              <thead>
                <tr>
                  <th>Ngày</th><th>Phòng ban</th><th>Hoàn thành</th><th>Chất lượng</th><th>Sự cố</th><th>Người nộp</th><th>Cập nhật</th>
                </tr>
              </thead>
              <tbody id="reportBody"></tbody>
            </table>
          </div>
        </section>

        <section class="card section app-page hidden" data-page="reports" id="reportsSection">
          <div id="reportsFolderView">
            <h3>Trung tâm Báo cáo</h3>
            <p class="muted" style="margin:4px 0 10px;">Chọn thư mục Báo cáo của từng bộ phận để xem chỉ số chi tiết theo ngày.</p>
            <div class="reports-folder-grid">
              <button class="reports-folder-card" type="button" data-report-dept="marketing">
                <strong>Báo cáo Marketing</strong>
                <span>Lead, tương tác và chuyển đổi theo ngày</span>
              </button>
              <button class="reports-folder-card" type="button" data-report-dept="telesale">
                <strong>Báo cáo Telesale</strong>
                <span>Lịch trải nghiệm và tỉ lệ đặt lịch</span>
              </button>
              <button class="reports-folder-card" type="button" data-report-dept="consultant">
                <strong>Báo cáo Tư vấn</strong>
                <span>Doanh số, số hợp đồng, tỉ lệ ký</span>
              </button>
              <button class="reports-folder-card" type="button" data-report-dept="nurse">
                <strong>Báo cáo Điều dưỡng</strong>
                <span>Khách đang chăm, hoàn tất và điểm dịch vụ</span>
              </button>
            </div>
          </div>

          <div id="reportsDetailView" class="hidden">
            <div class="reports-detail-head">
              <div>
                <h3 id="reportsDetailTitle">Báo cáo bộ phận</h3>
                <p class="muted" id="reportsDetailDesc" style="margin:3px 0 0;">Chi tiết KPI theo ngày</p>
              </div>
              <div class="reports-detail-actions">
                <button class="btn secondary" id="exportReportsExcelBtn" type="button">Xuất Excel (CSV)</button>
                <button class="btn warn" id="exportReportsPdfBtn" type="button">Xuất PDF</button>
                <button class="btn secondary" id="reportsSyncMetricsBtn" type="button">Đồng bộ sang Trang chỉ số</button>
                <button class="btn secondary" id="reportsBackBtn" type="button">← Quay về thư mục</button>
              </div>
            </div>
            <div class="reports-filter" style="margin-top:10px;">
              <div>
                <label>Từ ngày</label>
                <input id="reportsStartDate" type="date" />
              </div>
              <div>
                <label>Đến ngày</label>
                <input id="reportsEndDate" type="date" />
              </div>
              <div id="reportsMarketingFilterWrap" class="hidden">
                <label>Tên marketing</label>
                <select id="reportsMarketingFilter">
                  <option value="">Tất cả marketing</option>
                </select>
              </div>
              <div id="reportsConsultantFilterWrap" class="hidden">
                <label>Tên tư vấn</label>
                <select id="reportsConsultantFilter">
                  <option value="">Tất cả tư vấn</option>
                </select>
              </div>
              <div id="reportsTelesaleFilterWrap" class="hidden">
                <label>Tên telesale</label>
                <select id="reportsTelesaleFilter">
                  <option value="">Tất cả telesale</option>
                </select>
              </div>
              <button class="btn secondary" id="applyReportsFilterBtn" type="button">Áp dụng</button>
              <button class="btn secondary" id="resetReportsFilterBtn" type="button">Đặt lại tháng này</button>
              <div class="muted" id="reportsSummary" style="display:flex;align-items:center;font-size:0.83rem;">--</div>
            </div>
            <div class="reports-table-wrap" style="margin-top:10px;">
              <table id="reportsTable">
                <thead>
                  <tr>
                    <th id="reportsPrimaryCol">Ngày</th>
                    <th id="reportsMetricCol1">Chỉ số 1</th>
                    <th id="reportsMetricCol2">Chỉ số 2</th>
                    <th id="reportsMetricCol3">Chỉ số 3</th>
                    <th id="reportsMetricCol4">Chỉ số 4</th>
                  </tr>
                </thead>
                <tbody id="reportsBody"></tbody>
              </table>
            </div>
          </div>
        </section>

        <section id="adminUserSection" class="card section app-page hidden" data-page="hr">
          <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
            <h3>Danh sách nhân sự</h3>
            <div style="display:flex;gap:8px;flex-wrap:wrap;">
              <button class="btn secondary" id="openPermissionModalBtn" type="button">Phân quyền</button>
              <button class="btn secondary" id="openUserModalBtn" type="button">+ Thêm nhân sự</button>
            </div>
          </div>
          <span class="alert" id="userManageStatus"></span>
          <div class="customer-filter form-grid" style="margin-top:10px;">
            <div class="customer-search-field">
              <label>Tìm kiếm (Họ tên / Username)</label>
              <input id="hrQuickSearch" placeholder="Nhập tên hoặc username..." />
            </div>
            <div>
              <label>Vai trò</label>
              <select id="hrFilterRole">
                <option value="">Tất cả vai trò</option>
                <option value="staff">Nhân viên</option>
                <option value="head">Trưởng bộ phận</option>
                <option value="ceo">CEO</option>
                <option value="admin">Admin</option>
              </select>
            </div>
            <div>
              <label>Phòng ban</label>
              <select id="hrFilterDept">
                <option value="">Tất cả phòng ban</option>
                <option>Ban điều hành</option>
                ${DEPARTMENTS.map((d) => `<option>${d}</option>`).join("")}
              </select>
            </div>
            <div>
              <label>Trạng thái</label>
              <select id="hrFilterStatus">
                <option value="">Tất cả trạng thái</option>
                <option value="active">Đang hoạt động</option>
                <option value="suspended">Tạm dừng</option>
              </select>
            </div>
          </div>
          <div class="tables" style="margin-top:10px;">
            <table>
              <thead>
                <tr>
                  <th>#</th>
                  <th>Mã NV</th>
                  <th>Họ tên</th>
                  <th>Username</th>
                  <th>Vai trò</th>
                  <th>Phòng ban</th>
                  <th>Trạng thái</th>
                  <th>Ngày tạo</th>
                  <th>Thao tác</th>
                </tr>
              </thead>
              <tbody id="userBody"></tbody>
            </table>
          </div>
          <div class="customer-modal hidden" id="userModal">
            <div class="customer-modal-backdrop" id="userModalBackdrop"></div>
            <div class="customer-modal-panel card">
              <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;">
                <h3 id="userModalTitle">Thêm nhân sự mới</h3>
                <button class="btn warn" id="closeUserModalBtn" type="button">Đóng</button>
              </div>
              <div class="form-grid" style="grid-template-columns:repeat(2,minmax(0,1fr));margin-top:8px;">
                <div>
                  <label>Mã nhân viên</label>
                  <input id="userCode" placeholder="NR001" readonly style="background:#f1f5f9;cursor:default;color:#6b7280;font-weight:600;letter-spacing:0.5px;" />
                </div>
                <div>
                  <label>Họ tên</label>
                  <input id="userFullName" placeholder="Nguyễn Văn A" />
                </div>
                <div>
                  <label>Username</label>
                  <input id="userUsername" placeholder="user-name" />
                </div>
                <div>
                  <label>Mật khẩu</label>
                  <input id="userPassword" type="password" placeholder="Tối thiểu 6 ký tự" />
                </div>
                <div>
                  <label>Vai trò / Chức vụ</label>
                  <select id="userRoleKey">
                    <option value="staff">Nhân viên</option>
                    <option value="head">Trưởng bộ phận</option>
                    <option value="ceo">CEO</option>
                    <option value="admin">Admin</option>
                  </select>
                </div>
                <div style="grid-column:1/-1;">
                  <label>Phòng ban</label>
                  <select id="userDepartment">
                    <option>Ban điều hành</option>
                    ${DEPARTMENTS.map((d) => `<option>${d}</option>`).join("")}
                  </select>
                </div>
                <div>
                  <label>Số điện thoại</label>
                  <input id="userPhone" placeholder="0909 123 456" />
                </div>
                <div>
                  <label>Email</label>
                  <input id="userEmail" type="email" placeholder="nhanvien@email.com" />
                </div>
                <div>
                  <label>Địa chỉ</label>
                  <input id="userAddress" placeholder="Số nhà, đường, quận/huyện, tỉnh/thành" />
                </div>
                <div>
                  <label>STK ngân hàng</label>
                  <input id="userBankAccount" placeholder="Ngân hàng · Số tài khoản" />
                </div>
                <div style="grid-column:1/-1; display:flex; justify-content:flex-end; gap:8px; padding-top:4px;">
                  <button class="btn secondary" id="saveUserBtn" type="button">Tạo tài khoản</button>
                </div>
                <div style="grid-column:1/-1;">
                  <span class="alert" id="userModalStatus"></span>
                </div>
              </div>
            </div>
          </div>
          <div class="customer-modal hidden" id="hrProfileModal">
            <div class="customer-modal-backdrop" id="hrProfileModalBackdrop"></div>
            <div class="customer-modal-panel card" style="max-width:600px;">
              <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;">
                <div>
                  <h3 id="hrProfileTitle">Hồ sơ nhân sự</h3>
                  <p id="hrProfileSubtitle" class="muted" style="font-size:0.85rem;margin:2px 0 0;"></p>
                </div>
                <button class="btn warn" id="closeHrProfileBtn" type="button">Đóng</button>
              </div>
              <div id="hrProfileInfo" style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin:12px 0;background:#f8fafc;border-radius:10px;padding:12px;font-size:0.85rem;"></div>
              <div style="display:flex;justify-content:space-between;align-items:center;margin:10px 0 6px;">
                <strong style="font-size:0.9rem;">Tài liệu đính kèm</strong>
                <label class="btn secondary" style="cursor:pointer;font-size:0.8rem;padding:6px 12px;">
                  + Tải lên tài liệu
                  <input type="file" id="hrFileUpload" accept=".pdf,.doc,.docx,.jpg,.jpeg,.png,.xlsx,.xls" multiple style="display:none;" />
                </label>
              </div>
              <div id="hrProfileFileList" style="display:flex;flex-direction:column;gap:6px;max-height:280px;overflow-y:auto;"></div>
              <p class="alert" id="hrProfileStatus" style="margin-top:8px;"></p>
            </div>
          </div>

          <div class="customer-modal hidden" id="permissionModal">
            <div class="customer-modal-backdrop" id="permissionModalBackdrop"></div>
            <div class="customer-modal-panel card" style="max-width:920px;">
              <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
                <div>
                  <h3 style="margin:0;">Phân quyền theo vị trí</h3>
                  <p class="muted" style="margin:3px 0 0;font-size:0.84rem;">Thiết lập chức năng và trang được xem cho từng vai trò. Admin/CEO luôn có toàn quyền.</p>
                </div>
                <button class="btn warn" id="closePermissionModalBtn" type="button">Đóng</button>
              </div>
              <div class="form-grid" style="grid-template-columns:1fr;margin-top:10px;">
                <div>
                  <label>Vai trò</label>
                  <select id="permissionRoleSelect">
                    <option value="admin">Admin</option>
                    <option value="ceo">CEO</option>
                    <option value="head">Trưởng bộ phận</option>
                    <option value="staff">Nhân viên</option>
                  </select>
                </div>
                <div class="card" style="padding:10px;">
                  <h4 style="margin:0 0 8px;font-size:0.9rem;">Chức năng</h4>
                  <div id="permissionFeatureGrid" class="reports-folder-grid" style="grid-template-columns:repeat(3,minmax(0,1fr));margin-top:0;"></div>
                </div>
                <div class="card" style="padding:10px;">
                  <h4 style="margin:0 0 8px;font-size:0.9rem;">Giới hạn trang xem</h4>
                  <div id="permissionPageGrid" class="reports-folder-grid" style="grid-template-columns:repeat(3,minmax(0,1fr));margin-top:0;"></div>
                </div>
                <div style="display:flex;justify-content:flex-end;gap:8px;align-items:center;">
                  <span class="alert" id="permissionModalStatus" style="margin:0;flex:1;"></span>
                  <button class="btn secondary" id="resetPermissionBtn" type="button">Đặt lại mặc định</button>
                  <button class="btn secondary" id="savePermissionBtn" type="button">Lưu phân quyền</button>
                </div>
              </div>
            </div>
          </div>
        </section>

        <section class="card section app-page hidden" data-page="customers" id="customerSection">
          <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
            <h3>Quản lí khách hàng</h3>
            <button class="btn secondary" id="openCustomerModalBtn" type="button">Thêm khách mới</button>
          </div>
          <span class="alert" id="customerStatusMessage"></span>
          <div class="customer-filter form-grid">
            <div class="customer-search-field">
              <label>Tìm nhanh (Tên / Người liên hệ / SĐT / Email)</label>
              <input id="customerQuickSearch" placeholder="Nhập từ khóa để lọc nhanh" />
            </div>
            <div class="customer-date-field" id="customerDateRangeField">
              <label>Khoảng ngày cập nhật</label>
              <button class="date-range-trigger" id="customerDateRangeTrigger" type="button">Mọi ngày</button>
              <div class="date-range-popover hidden" id="customerDateRangePopover">
                <div class="date-range-presets" id="customerDateRangePresets">
                  <button class="btn secondary" type="button" data-range-preset="today">Hôm nay</button>
                  <button class="btn secondary" type="button" data-range-preset="7d">7 ngày</button>
                  <button class="btn secondary" type="button" data-range-preset="thisMonth">Tháng này</button>
                  <button class="btn secondary" type="button" data-range-preset="lastMonth">Tháng trước</button>
                </div>
                <div class="date-range-row">
                  <div>
                    <label>Từ ngày</label>
                    <input id="customerFilterStartDate" type="date" />
                  </div>
                  <div>
                    <label>Đến ngày</label>
                    <input id="customerFilterEndDate" type="date" />
                  </div>
                </div>
                <div class="date-range-actions">
                  <button class="btn secondary" id="applyCustomerDateRangeBtn" type="button">Xong</button>
                  <button class="btn secondary" id="clearCustomerDateRangeBtn" type="button">Xóa ngày</button>
                </div>
              </div>
            </div>
            <div>
              <label>Người phụ trách</label>
              <select id="customerFilterOwner"></select>
            </div>
            <div>
              <label>Trạng thái chăm sóc</label>
              <select id="customerFilterStatus"></select>
            </div>
            <div>
              <label>Nguồn data</label>
              <select id="customerFilterSource"></select>
            </div>
            <button class="btn secondary" id="applyCustomerFilterBtn" type="button">Áp dụng bộ lọc</button>
            <button class="btn secondary" id="resetCustomerFilterBtn" type="button">Đặt lại bộ lọc</button>
            <button class="btn secondary" id="exportCustomerExcelBtn" type="button">Xuất Excel (CSV)</button>
            <button class="btn warn" id="exportCustomerPdfBtn" type="button">Xuất PDF khách hàng</button>
          </div>
          <div class="tables" style="margin-top:10px;">
            <table>
              <thead>
                <tr><th>Khách hàng</th><th>Liên hệ</th><th>Phân loại</th><th>Người phụ trách</th><th>Trạng thái chăm sóc</th><th>Nguồn data</th><th>Ghi chú</th><th>Cập nhật</th><th>Thao tác</th></tr>
              </thead>
              <tbody id="customerBody"></tbody>
            </table>
          </div>

          <div class="customer-modal hidden" id="customerModal">
            <div class="customer-modal-backdrop" id="customerModalBackdrop"></div>
            <div class="customer-modal-panel card">
              <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;">
                <h3 id="customerModalTitle">Thêm khách hàng mới</h3>
                <button class="btn warn" id="closeCustomerModalBtn" type="button">Đóng</button>
              </div>
              <div class="form-grid" style="grid-template-columns:repeat(2,minmax(0,1fr));margin-top:8px;">
                <div>
                  <label>Tên khách hàng</label>
                  <input id="customerName" placeholder="Công ty ABC" />
                </div>
                <div>
                  <label>Người liên hệ</label>
                  <input id="customerContactPerson" placeholder="Nguyễn Văn A" />
                </div>
                <div>
                  <label>Số điện thoại</label>
                  <input id="customerPhone" placeholder="09xxxxxxxx" />
                </div>
                <div>
                  <label>Email</label>
                  <input id="customerEmail" placeholder="abc@company.com" />
                </div>
                <div style="grid-column:1 / -1;">
                  <label>Địa chỉ</label>
                  <input id="customerAddress" placeholder="Địa chỉ (tự điền hoặc chọn bên dưới)" autocomplete="off" />
                </div>
                <div style="grid-column:1 / -1; display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:10px;">
                  <div>
                    <label>Tỉnh/Thành phố</label>
                    <select id="customerAddrProvince">
                      <option value="">Lựa chọn</option>
                      ${Object.keys(VN_ADDRESSES).map((p) => `<option value="${p}">${p}</option>`).join("")}
                    </select>
                  </div>
                  <div>
                    <label>Quận/Huyện</label>
                    <select id="customerAddrDistrict" disabled>
                      <option value="">Lựa chọn</option>
                    </select>
                  </div>
                  <div>
                    <label>Phường/Xã</label>
                    <select id="customerAddrWard" disabled>
                      <option value="">Lựa chọn</option>
                    </select>
                  </div>
                </div>
                <div>
                  <label>Phân loại</label>
                  <select id="customerTier">
                    <option>Tiêu chuẩn</option>
                    <option>Tiềm năng</option>
                    <option>VIP</option>
                  </select>
                </div>
                <div>
                  <label>Trạng thái</label>
                  <select id="customerStatus">
                    ${CUSTOMER_STATUSES.map((status) => `<option>${status}</option>`).join("")}
                  </select>
                </div>
                <div>
                  <label>Người phụ trách</label>
                  <select id="customerOwner"></select>
                </div>
                <div>
                  <label>Nguồn data</label>
                  <select id="customerSource">
                    ${CUSTOMER_SOURCES.map((source) => `<option>${source}</option>`).join("")}
                  </select>
                </div>
                <div style="grid-column:1 / -1;">
                  <label>Nhu cầu</label>
                  <input id="customerDemand" placeholder="Nhu cầu/chính sách khách quan tâm" />
                </div>
                <div style="grid-column:1 / -1;">
                  <label>Ghi chú chăm sóc</label>
                  <textarea id="customerNote" rows="3" placeholder="Ghi đầy đủ thông tin chăm sóc: lịch gọi, phản hồi, bước tiếp theo..."></textarea>
                </div>
                <div style="grid-column:1 / -1;">
                  <button class="btn secondary" id="saveCustomerBtn" type="button">Lưu khách hàng</button>
                </div>
              </div>
            </div>
          </div>
        </section>

        <section class="card section app-page hidden" data-page="inventory" id="inventorySection">
          <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
            <div>
              <h3>Quản lí kho vật tư</h3>
              <p class="muted" style="margin:2px 0 0;">Theo dõi tồn kho, cảnh báo gần hết và thao tác xuất/nhập theo từng sản phẩm.</p>
            </div>
            <div style="display:flex;gap:8px;flex-wrap:wrap;">
              <button class="btn secondary" id="openInventoryModalBtn" type="button">+ Thêm vật tư mới</button>
              <button class="btn secondary" id="openInventoryHistoryBtn" type="button">📋 Lịch sử &amp; Thống kê</button>
              <button class="btn secondary" id="exportInventoryExcelBtn" type="button">Xuất Excel (CSV)</button>
              <button class="btn warn" id="exportInventoryPdfBtn" type="button">Xuất PDF kho</button>
            </div>
          </div>
          <div class="customer-filter form-grid" style="margin-top:10px;">
            <div class="customer-search-field">
              <label>Tìm kiếm sản phẩm</label>
              <input id="inventorySearch" placeholder="Mã sản phẩm / Tên sản phẩm..." />
            </div>
            <div>
              <label>Trạng thái sản phẩm</label>
              <select id="inventoryFilterStatus">
                <option value="">Tất cả trạng thái</option>
                <option value="active">Đang kinh doanh</option>
                <option value="inactive">Ngừng kinh doanh</option>
              </select>
            </div>
            <div>
              <label>Tồn kho</label>
              <select id="inventoryFilterStock">
                <option value="">Tất cả</option>
                <option value="safe">An toàn</option>
                <option value="low">Gần hết</option>
                <option value="out">Hết hàng</option>
              </select>
            </div>
          </div>
          <div style="margin-top:10px;display:flex;align-items:center;gap:16px;flex-wrap:wrap;">
            <div class="muted" id="inventoryFilterSummary" style="font-size:0.83rem;"></div>
            <div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:6px 16px;font-size:0.88rem;">
              💰 Tổng giá trị tồn kho: <strong id="inventoryTotalValue">0 đ</strong>
            </div>
          </div>
          <span class="alert" id="inventoryStatusMessage"></span>

          <div class="tables" style="margin-top:10px;">
            <table>
              <thead>
                <tr>
                  <th>#</th>
                  <th>Mã sản phẩm</th>
                  <th>Tên sản phẩm</th>
                  <th>Giá nhập</th>
                  <th>Giá bán</th>
                  <th>Số lượng</th>
                  <th>Nhà cung cấp</th>
                  <th>Hạn sử dụng</th>
                  <th>Giá trị tồn</th>
                  <th>Trạng thái</th>
                  <th>Cảnh báo</th>
                  <th>Cập nhật</th>
                  <th>Thao tác</th>
                </tr>
              </thead>
              <tbody id="inventoryBody"></tbody>
            </table>
          </div>

          <div class="customer-modal hidden" id="inventoryHistoryModal">
            <div class="customer-modal-backdrop" id="inventoryHistoryModalBackdrop"></div>
            <div class="customer-modal-panel card" style="max-width:1100px;width:96vw;max-height:90vh;overflow-y:auto;">
              <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:12px;">
                <h3 style="margin:0;">📋 Lịch sử &amp; Thống kê xuất nhập kho</h3>
                <button class="btn warn" id="closeInventoryHistoryBtn" type="button">Đóng</button>
              </div>
              <div class="customer-filter form-grid" style="grid-template-columns:repeat(4,minmax(0,1fr));">
                <div>
                  <label>Thống kê từ ngày</label>
                  <input id="inventoryStatsStart" type="date" />
                </div>
                <div>
                  <label>Đến ngày</label>
                  <input id="inventoryStatsEnd" type="date" />
                </div>
                <div style="display:flex;align-items:flex-end;gap:8px;">
                  <button class="btn secondary" id="applyInventoryStatsBtn" type="button">Áp dụng</button>
                  <button class="btn secondary" id="resetInventoryStatsBtn" type="button">Đặt lại</button>
                </div>
                <div class="muted" id="inventoryStatsRangeText" style="display:flex;align-items:flex-end;font-size:0.82rem;"></div>
              </div>
              <div class="kpis" style="margin-top:10px;">
                <article class="kpi"><div class="label">Tổng nhập</div><div class="value" id="inventoryInTotal">0</div></article>
                <article class="kpi"><div class="label">Tổng xuất</div><div class="value" id="inventoryOutTotal">0</div></article>
                <article class="kpi"><div class="label">Giao dịch</div><div class="value" id="inventoryTxnTotal">0</div></article>
                <article class="kpi"><div class="label">Chênh lệch (Nhập - Xuất)</div><div class="value" id="inventoryNetTotal">0</div></article>
              </div>
              <div class="tables" style="margin-top:12px;">
                <table>
                  <thead>
                    <tr>
                      <th>Thời gian</th>
                      <th>Mã sản phẩm</th>
                      <th>Tên sản phẩm</th>
                      <th>Loại</th>
                      <th>Số lượng</th>
                      <th>Tồn sau giao dịch</th>
                      <th>Người thao tác</th>
                      <th>Ghi chú</th>
                    </tr>
                  </thead>
                  <tbody id="inventoryHistoryBody"></tbody>
                </table>
              </div>
            </div>
          </div>

          <div class="customer-modal hidden" id="inventoryModal">
            <div class="customer-modal-backdrop" id="inventoryModalBackdrop"></div>
            <div class="customer-modal-panel card" style="max-width:760px;">
              <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;">
                <h3 id="inventoryModalTitle">Thêm vật tư mới</h3>
                <button class="btn warn" id="closeInventoryModalBtn" type="button">Đóng</button>
              </div>
              <div class="form-grid" style="grid-template-columns:repeat(2,minmax(0,1fr));margin-top:8px;">
                <div>
                  <label>Mã sản phẩm</label>
                  <input id="inventoryProductCode" placeholder="VT-001" />
                </div>
                <div>
                  <label>Tên sản phẩm</label>
                  <input id="inventoryProductName" placeholder="Tên vật tư" />
                </div>
                <div>
                  <label>Giá nhập</label>
                  <input id="inventoryPurchasePrice" type="number" min="0" step="1000" placeholder="0" />
                </div>
                <div>
                  <label>Giá bán</label>
                  <input id="inventorySalePrice" type="number" min="0" step="1000" placeholder="0" />
                </div>
                <div>
                  <label>Số lượng ban đầu</label>
                  <input id="inventoryQuantity" type="number" min="0" step="1" placeholder="0" />
                </div>
                <div>
                  <label>Ngưỡng cảnh báo gần hết</label>
                  <input id="inventoryAlertThreshold" type="number" min="0" step="1" placeholder="20" />
                </div>
                <div>
                  <label>Nhà cung cấp</label>
                  <input id="inventorySupplier" placeholder="Tên nhà cung cấp" />
                </div>
                <div>
                  <label>Hạn sử dụng</label>
                  <input id="inventoryExpiryDate" type="date" />
                </div>
                <div>
                  <label>Trạng thái</label>
                  <select id="inventoryProductStatus">
                    <option value="active">Đang kinh doanh</option>
                    <option value="inactive">Ngừng kinh doanh</option>
                  </select>
                </div>
                <div style="display:flex;align-items:flex-end;justify-content:flex-end;">
                  <button class="btn secondary" id="saveInventoryBtn" type="button">Lưu vật tư</button>
                </div>
                <div style="grid-column:1/-1;">
                  <span class="alert" id="inventoryModalStatus"></span>
                </div>
              </div>
            </div>
          </div>

          <div class="customer-modal hidden" id="inventoryTxnModal">
            <div class="customer-modal-backdrop" id="inventoryTxnModalBackdrop"></div>
            <div class="customer-modal-panel card" style="max-width:520px;">
              <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;">
                <h3 id="inventoryTxnTitle">Xuất/Nhập kho</h3>
                <button class="btn warn" id="closeInventoryTxnModalBtn" type="button">Đóng</button>
              </div>
              <div class="form-grid" style="grid-template-columns:1fr;margin-top:8px;">
                <div>
                  <label>Sản phẩm</label>
                  <input id="inventoryTxnProduct" readonly style="background:#f1f5f9;color:#64748b;" />
                </div>
                <div>
                  <label>Loại giao dịch</label>
                  <select id="inventoryTxnType">
                    <option value="in">Nhập kho</option>
                    <option value="out">Xuất kho</option>
                  </select>
                </div>
                <div>
                  <label>Số lượng</label>
                  <input id="inventoryTxnQty" type="number" min="1" step="1" placeholder="Nhập số lượng" />
                </div>
                <div>
                  <label>Ghi chú</label>
                  <input id="inventoryTxnNote" placeholder="Ví dụ: Nhập từ NCC A / Xuất cho đơn hàng B" />
                </div>
                <div style="display:flex;justify-content:flex-end;">
                  <button class="btn secondary" id="saveInventoryTxnBtn" type="button">Xác nhận</button>
                </div>
                <div>
                  <span class="alert" id="inventoryTxnStatus"></span>
                </div>
              </div>
            </div>
          </div>
        </section>

        <section class="card section app-page hidden" data-page="workflow" id="workflowSection">
          <!-- Hidden legacy form kept for JS compatibility -->
          <form id="reportForm" style="display:none;">
            <input id="reportDate" type="date" />
            <select id="department">${DEPARTMENTS.map((d) => `<option>${d}</option>`).join("")}</select>
            <input id="submitter" /><input id="completion" type="number" /><input id="quality" type="number" /><input id="issues" type="number" />
            <button id="submitReportBtn" type="submit"></button>
          </form>
          <div class="alert" id="submitStatus" style="display:none;"></div>

          <div class="card" style="margin-bottom:14px;padding:12px;">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
              <div>
                <h4 style="margin:0 0 4px;">Danh mục quy trình theo phòng ban</h4>
                <p class="muted" style="margin:0;font-size:0.83rem;">Chọn phòng ban để xem trang quy trình chi tiết riêng.</p>
              </div>
              <div style="display:flex;gap:8px;flex-wrap:wrap;">
                <button class="btn secondary" type="button" data-workflow-target="hcns">HCNS</button>
                <button class="btn secondary" type="button" data-workflow-target="marketing">Marketing</button>
                <button class="btn secondary" type="button" data-workflow-target="kinh-doanh">Kinh doanh</button>
                <button class="btn secondary" type="button" data-workflow-target="tu-van">Tư vấn</button>
                <button class="btn secondary" type="button" data-workflow-target="dieu-hanh-ca">Điều hành ca</button>
                <button class="btn secondary" type="button" data-workflow-target="dieu-duong">Điều dưỡng</button>
                <button class="btn secondary" type="button" data-workflow-target="cskh">CSKH</button>
                <button class="btn secondary" type="button" data-workflow-target="ke-toan">Kế toán</button>
                <button class="btn secondary" type="button" data-workflow-target="giam-sat">Giám sát</button>
              </div>
            </div>
          </div>

          <div id="workflowSystemView">

          <!-- Page header -->
          <div style="text-align:center;margin-bottom:20px;">
            <h3 style="font-size:1.3rem;font-weight:800;margin:0 0 4px;letter-spacing:0.5px;">SƠ ĐỒ QUY TRÌNH LÀM VIỆC</h3>
            <p class="muted" style="margin:0;font-size:0.85rem;">Quy trình vận hành toàn hệ thống theo từng phòng ban.</p>
          </div>

          <!-- Top: Director -->
          <div style="text-align:center;margin-bottom:8px;">
            <div style="display:inline-flex;align-items:center;gap:8px;background:#1a1a2e;color:#fff;padding:10px 32px;border-radius:8px;font-weight:700;font-size:0.95rem;letter-spacing:0.5px;">👤 GIÁM ĐỐC VẬN HÀNH</div>
          </div>
          <div style="text-align:center;font-size:1.4rem;color:#bbb;line-height:1.2;margin-bottom:8px;">↓</div>

          <!-- 3-column layout -->
          <div style="display:grid;grid-template-columns:210px minmax(0,1fr) 210px;gap:14px;align-items:start;">

            <!-- LEFT: PHÒNG HCNS -->
            <aside style="background:#fffde7;border:1.5px solid #fbc02d;border-radius:10px;overflow:hidden;">
              <div style="background:#f9a825;color:#333;font-weight:700;padding:8px 10px;font-size:0.82rem;text-align:center;">📁 PHÒNG HCNS</div>
              <ul style="margin:0;padding:10px 10px 10px 24px;font-size:0.76rem;line-height:1.8;color:#444;list-style:disc;">
                <li>Xây dựng cơ cấu tổ chức và mô tả công việc, xây dựng JD</li>
                <li>Lên kế hoạch và triển khai tuyển dụng, đào tạo KPI, Deadline</li>
                <li>Sàng lọc hồ sơ, phỏng vấn và đánh giá ứng viên</li>
                <li>Quản lý lương và phúc lợi, chính sách lương thưởng</li>
                <li>Hoàn tất và lưu giữ hồ sơ, chứng từ nhân sự</li>
                <li>Xây dựng và cập nhật quy chế nội bộ, nội quy nhân sự</li>
                <li>Chăm sóc và quản lý công tác nhân viên nội bộ</li>
                <li>Giải quyết quyền lợi nhân viên (nghỉ phép, KPI)</li>
                <li>Xây dựng văn hóa doanh nghiệp và gắn kết nội bộ</li>
                <li>Phối hợp xử lý hồ sơ lao động và các chế độ</li>
              </ul>
            </aside>

            <!-- CENTER: main flow -->
            <div style="display:flex;flex-direction:column;gap:0;">

              <!-- Marketing -->
              <div style="border-radius:10px;overflow:hidden;border:1.5px solid #2d6a4f;">
                <div style="background:#2d6a4f;color:#fff;font-weight:700;padding:8px 14px;font-size:0.85rem;text-align:center;">📣 PHÒNG MARKETING</div>
                <ul style="margin:0;padding:8px 12px 8px 28px;font-size:0.77rem;line-height:1.75;color:#333;background:#f0fff4;list-style:disc;">
                  <li>Phân tích thị trường, nghiên cứu khách hàng và xu thế dịch vụ</li>
                  <li>Lên kế hoạch và xúc tiến chiến dịch Marketing</li>
                  <li>Triển khai các chiến dịch theo kênh online và offline</li>
                  <li>Thu thập và xử lý phản hồi, điều chỉnh chiến lược phù hợp</li>
                  <li>Phối hợp với phòng Kinh Doanh và nhận feedback</li>
                </ul>
              </div>
              <div style="text-align:center;font-size:1.3rem;color:#bbb;line-height:1.3;">↓</div>

              <!-- Kinh Doanh -->
              <div style="border-radius:10px;overflow:hidden;border:1.5px solid #1b4332;">
                <div style="background:#1b4332;color:#fff;font-weight:700;padding:8px 14px;font-size:0.85rem;text-align:center;">💼 PHÒNG KINH DOANH</div>
                <ul style="margin:0;padding:8px 12px 8px 28px;font-size:0.77rem;line-height:1.75;color:#333;background:#d8f3dc;list-style:disc;">
                  <li>Tiếp nhận Lead từ bộ phận Marketing, ưu tiên CRM khách hàng</li>
                  <li>Lên kịch bản và thực hiện tư vấn cho khách hàng</li>
                  <li>Chốt sale, đặt lịch và theo dõi khách hàng chưa tương tác</li>
                  <li>Lắng nghe và xử lý phản hồi, phối hợp Marketing nâng cao chất lượng</li>
                  <li>Kết hợp với Điều Hành Ca để tiếp tục trải nghiệm khách hàng</li>
                </ul>
              </div>
              <div style="text-align:center;font-size:1.3rem;color:#bbb;line-height:1.3;">↓</div>

              <!-- Row: Tư Vấn | Điều Hành Ca | Điều Dưỡng -->
              <div style="display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;">

                <div style="border-radius:10px;overflow:hidden;border:1.5px solid #d97706;">
                  <div style="background:#d97706;color:#fff;font-weight:700;padding:8px 8px;font-size:0.8rem;text-align:center;">🗣️ TƯ VẤN</div>
                  <ul style="margin:0;padding:8px 8px 8px 20px;font-size:0.73rem;line-height:1.75;color:#333;background:#fff8e1;list-style:disc;">
                    <li>Tiếp nhận lịch và thông tin Lead, đăng ký tư vấn</li>
                    <li>Tư vấn dịch vụ, phác đồ và phương án phù hợp</li>
                    <li>Cam kết quyền lợi, lộ trình và chi phí dịch vụ</li>
                    <li>Xử lý thắc mắc, giải đáp và chốt khách hàng</li>
                    <li>Trực tiếp thu tiền và theo dõi hậu điểm</li>
                    <li>Chuyển tiếp theo dõi trải nghiệm sau dịch vụ</li>
                  </ul>
                </div>

                <div style="border-radius:10px;overflow:hidden;border:1.5px solid #92400e;">
                  <div style="background:#92400e;color:#fff;font-weight:700;padding:8px 8px;font-size:0.8rem;text-align:center;">⚙️ ĐIỀU HÀNH CA</div>
                  <ul style="margin:0;padding:8px 8px 8px 20px;font-size:0.73rem;line-height:1.75;color:#333;background:#fef3c7;list-style:disc;">
                    <li>Quản lý và điều phối lịch hàng ngày</li>
                    <li>Phối hợp với Tư Vấn, Kinh Doanh và Điều Dưỡng</li>
                    <li>Điều tiết lịch và tổng hợp kết quả mỗi ngày</li>
                    <li>Đảm bảo thông tin chính xác giữa các bộ phận</li>
                    <li>Báo cáo và phản hồi về giám đốc vận hành</li>
                  </ul>
                </div>

                <div style="border-radius:10px;overflow:hidden;border:1.5px solid #b91c1c;">
                  <div style="background:#b91c1c;color:#fff;font-weight:700;padding:8px 8px;font-size:0.8rem;text-align:center;">💉 ĐIỀU DƯỠNG</div>
                  <ul style="margin:0;padding:8px 8px 8px 20px;font-size:0.73rem;line-height:1.75;color:#333;background:#fef2f2;list-style:disc;">
                    <li>Nhận ca từ bộ phận ĐHC với mẫu đăng ký và lịch</li>
                    <li>Thực hiện chăm sóc theo quy trình chuyên môn</li>
                    <li>Báo cáo check in, check out về bộ phận ĐHC</li>
                    <li>Ghi nhận phản hồi và nhắc nhở sau dịch vụ</li>
                    <li>Chuẩn bị và đảm bảo chất lượng sản phẩm chăm sóc</li>
                    <li>Báo cáo kết thúc ca và tổng hợp về công ty</li>
                  </ul>
                </div>
              </div>
              <div style="text-align:center;font-size:1.3rem;color:#bbb;line-height:1.3;">↓</div>

              <!-- Chăm Sóc Khách Hàng -->
              <div style="border-radius:10px;overflow:hidden;border:1.5px solid #be123c;">
                <div style="background:#be123c;color:#fff;font-weight:700;padding:8px 14px;font-size:0.85rem;text-align:center;">💬 CHĂM SÓC KHÁCH HÀNG</div>
                <ul style="margin:0;padding:8px 12px 8px 28px;font-size:0.77rem;line-height:1.75;color:#333;background:#fff1f2;list-style:disc;">
                  <li>Tiếp nhận và theo dõi khách hàng sau dịch vụ theo giao thức hoàn thiện</li>
                  <li>Theo dõi, nhắc lịch tái khám và đặt lịch hẹn quay lại</li>
                  <li>Giao tiếp và kết nối giữa các phòng ban liên quan</li>
                  <li>Điều phối lịch trong ca và đảm bảo chất lượng dịch vụ</li>
                  <li>Nắm bắt nhu cầu, cảm nhận và đặt lịch mở rộng bền lâu</li>
                </ul>
              </div>
              <div style="text-align:center;font-size:1.3rem;color:#bbb;line-height:1.3;">↓</div>

              <!-- Kế Toán -->
              <div style="border-radius:10px;overflow:hidden;border:1.5px solid #1d4ed8;">
                <div style="background:#1d4ed8;color:#fff;font-weight:700;padding:8px 14px;font-size:0.85rem;text-align:center;">🧾 PHÒNG KẾ TOÁN</div>
                <ul style="margin:0;padding:8px 12px 8px 28px;font-size:0.77rem;line-height:1.75;color:#333;background:#eff6ff;list-style:disc;">
                  <li>Theo dõi doanh thu hàng ngày, hàng tuần và hàng tháng</li>
                  <li>Kiểm tra, đối soát thanh toán khách hàng và thu chi nội bộ</li>
                  <li>Theo dõi các khoản chi phí nội bộ</li>
                  <li>Lập phiếu thu, phiếu chi, hợp đồng và chứng từ</li>
                  <li>Xuất hoá đơn VAT chi tiết từng lần sử dụng</li>
                  <li>Lập kế hoạch tài chính và cập nhật công việc thực hiện</li>
                  <li>Theo dõi công nợ và doanh thu từng điều dưỡng</li>
                  <li>Quản lý thiết bị, tài sản và các khoản liên quan</li>
                  <li>Kiểm soát hàng tồn kho, vật tư, đồng phục và sử dụng</li>
                </ul>
              </div>
            </div><!-- end center flow -->

            <!-- RIGHT: GIÁM SÁT -->
            <aside style="background:#f0fdf4;border:1.5px solid #4ade80;border-radius:10px;overflow:hidden;">
              <div style="background:#16a34a;color:#fff;font-weight:700;padding:8px 10px;font-size:0.82rem;text-align:center;">🔍 GIÁM SÁT</div>
              <ul style="margin:0;padding:10px 10px 10px 24px;font-size:0.76rem;line-height:1.8;color:#444;list-style:disc;">
                <li>Kiểm tra triển khai hoạt động của các phòng ban</li>
                <li>Theo dõi kết quả, đánh giá và đề xuất đào tạo cải thiện</li>
                <li>Kiểm tra văn phòng, đôn đốc báo cáo và tuân thủ quy trình</li>
                <li>Kiểm tra chất lượng điều dưỡng và dịch vụ</li>
                <li>Nhắc nhở phân phối dịch vụ và bán sản phẩm</li>
                <li>Giám sát chất lượng dịch vụ và xử lý tình huống phát sinh</li>
                <li>Kiểm tra và chứng thực dịch vụ, tiếp nhận phản hồi</li>
                <li>Đảm bảo quy trình và thực hiện KPI chất lượng</li>
                <li>Giám sát đào tạo, quy định và chính sách công ty</li>
                <li>Phân công đào tạo cho nhân viên mới</li>
                <li>Phân lịch trực để đảm bảo hoạt động mọi phòng ban</li>
                <li>Kiểm tra và phân phối chính sách đến từng phòng ban</li>
              </ul>
            </aside>

          </div><!-- end 3-col grid -->
          </div>

          <div id="workflowDetailView" class="hidden">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:10px;">
              <div>
                <h3 id="workflowDetailTitle" style="margin:0;">Chi tiết quy trình</h3>
                <p class="muted" style="margin:4px 0 0;font-size:0.84rem;">Các bước thực thi chuẩn theo phòng ban.</p>
              </div>
              <button class="btn secondary" id="workflowBackBtn" type="button">← Quay lại sơ đồ tổng thể</button>
            </div>
            <div class="card" style="padding:14px;">
              <ol id="workflowDetailSteps" style="margin:0;padding-left:22px;line-height:1.9;"></ol>
            </div>
          </div>
        </section>

        <section class="card section app-page hidden" data-page="policy" id="policySection">
          <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
            <div>
              <h3 style="margin:0;">Nội quy và cơ chế</h3>
              <p class="muted" style="margin:4px 0 0;font-size:0.84rem;">Chọn thư mục mẹ để mở danh sách thư mục con chi tiết theo từng phòng ban.</p>
            </div>
          </div>
          <div id="policyFolderRoot" style="margin-top:12px;"></div>
        </section>

        <section class="card section app-page hidden" data-page="schedule" id="scheduleSection">
          <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
            <div>
              <h3>Lịch trải nghiệm khách hàng</h3>
              <p class="muted" style="margin:2px 0 0;">Danh sách lịch đặt dịch vụ theo ngày, theo dõi tư vấn viên, điều dưỡng, trạng thái và kết quả hợp đồng.</p>
            </div>
            <div style="display:flex;gap:8px;flex-wrap:wrap;">
              <button class="btn secondary" id="openScheduleModalBtn" type="button">+ Thêm lịch</button>
              <button class="btn secondary" id="exportScheduleExcelBtn" type="button">Xuất Excel (CSV)</button>
              <button class="btn warn" id="exportSchedulePdfBtn" type="button">Xuất PDF</button>
            </div>
          </div>
          <div class="customer-filter form-grid" style="margin-top:10px;">
            <div class="customer-search-field">
              <label>Tìm nhanh (Tên / SĐT / Địa chỉ)</label>
              <input id="scheduleSearch" placeholder="Nhập từ khóa..." />
            </div>
            <div>
              <label>Tháng</label>
              <input id="scheduleFilterMonth" type="month" />
            </div>
            <div>
              <label>Trạng thái</label>
              <select id="scheduleFilterStatus">
                <option value="">Tất cả trạng thái</option>
                <option value="pending">Chờ xác nhận</option>
                <option value="confirmed">Đã xác nhận</option>
                <option value="completed">Đã hoàn thành</option>
                <option value="cancelled">Đã hủy</option>
              </select>
            </div>
            <div>
              <label>Tư vấn viên / Điều dưỡng / Sale</label>
              <select id="scheduleFilterStaff"></select>
            </div>
            <div>
              <label>Nguồn khách</label>
              <input id="scheduleFilterSource" placeholder="Facebook, Zalo..." />
            </div>
            <button class="btn secondary" id="applyScheduleFilterBtn" type="button">Lọc</button>
            <button class="btn secondary" id="resetScheduleFilterBtn" type="button">Đặt lại</button>
            <div class="muted" id="scheduleFilterSummary" style="display:flex;align-items:center;font-size:0.83rem;">Hiển thị tất cả lịch</div>
          </div>
          <div style="display:flex;gap:14px;margin-top:8px;flex-wrap:wrap;">
            <span class="schedule-legend confirmed">Đã xác nhận</span>
            <span class="schedule-legend completed">Hoàn thành</span>
            <span class="schedule-legend cancelled">Đã hủy</span>
            <span class="schedule-legend pending">Chờ xác nhận</span>
          </div>
          <div class="tables schedule-table-wrap" id="scheduleTableWrap" tabindex="0" style="margin-top:10px;overflow-x:hidden;">
            <table id="scheduleTable" style="width:100%;table-layout:fixed;">
              <thead>
                <tr>
                  <th>STT</th>
                  <th>Ngày trải nghiệm</th>
                  <th>Giờ trải nghiệm</th>
                  <th>Họ tên</th>
                  <th>Số điện thoại</th>
                  <th>Địa chỉ</th>
                  <th>Tình trạng</th>
                  <th>Ghi chú</th>
                  <th>Dịch vụ đăng ký</th>
                  <th>Tư vấn</th>
                  <th>Điều dưỡng</th>
                  <th>Nguồn data</th>
                  <th>Giá trị hợp đồng</th>
                  <th>Telesale</th>
                </tr>
              </thead>
              <tbody id="scheduleBody"></tbody>
            </table>
          </div>
          <div class="schedule-scroll-controls" id="scheduleScrollControls" aria-label="Điều khiển cuộn ngang danh sách lịch">
            <button class="schedule-scroll-btn" id="scheduleScrollLeftBtn" type="button" title="Cuộn sang trái">◀</button>
            <div class="schedule-bottom-scrollbar" id="scheduleBottomScroller" aria-label="Thanh cuộn ngang danh sách lịch">
              <div class="schedule-bottom-scrollbar-inner" id="scheduleBottomScrollerInner"></div>
            </div>
            <button class="schedule-scroll-btn" id="scheduleScrollRightBtn" type="button" title="Cuộn sang phải">▶</button>
          </div>

          <div class="customer-modal hidden" id="scheduleModal">
            <div class="customer-modal-backdrop" id="scheduleModalBackdrop"></div>
            <div class="customer-modal-panel card" style="max-width:960px;width:96vw;max-height:92vh;overflow-y:auto;">
              <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:12px;">
                <h3 id="scheduleModalTitle">Thêm lịch trải nghiệm</h3>
                <button class="btn warn" id="closeScheduleModalBtn" type="button">Đóng</button>
              </div>
              <div class="form-grid" style="grid-template-columns:repeat(3,minmax(0,1fr));">
                <div>
                  <label>Ngày đăng ký TN</label>
                  <input id="scheduleRegDate" type="date" />
                </div>
                <div>
                  <label>Giờ trải nghiệm</label>
                  <input id="scheduleTime" placeholder="10h, 14h30..." />
                </div>
                <div>
                  <label>Trạng thái</label>
                  <select id="scheduleStatus">
                    <option value="pending">Chờ xác nhận</option>
                    <option value="confirmed">Đã xác nhận</option>
                    <option value="completed">Đã hoàn thành</option>
                    <option value="cancelled">Đã hủy</option>
                  </select>
                </div>
                <div>
                  <label>Họ tên khách hàng</label>
                  <input id="scheduleName" placeholder="Nguyễn Thị A" />
                </div>
                <div>
                  <label>Số điện thoại</label>
                  <input id="schedulePhone" placeholder="09xxxxxxxx" />
                </div>
                <div>
                  <label>Tuổi mẹ</label>
                  <input id="scheduleMotherAge" placeholder="38 tuổi..." />
                </div>
                <div style="grid-column:1/-1;">
                  <label>Địa chỉ</label>
                  <input id="scheduleAddress" placeholder="Địa chỉ khách hàng" />
                </div>
                <div>
                  <label>Lịch sử sinh</label>
                  <input id="scheduleBirthHistory" placeholder="Con so 1, lần 2..." />
                </div>
                <div>
                  <label>Ngày sinh bé / Ưu tiên</label>
                  <input id="scheduleBabyBirthday" placeholder="10 ngày tuổi, 3/2026..." />
                </div>
                <div>
                  <label>Giai đoạn</label>
                  <input id="scheduleStage" placeholder="Mẹ bầu 20w, Em bé..." />
                </div>
                <div style="grid-column:1/-1;">
                  <label>Dịch vụ</label>
                  <input id="scheduleService" placeholder="Massage mẹ bầu, Chăm sóc mẹ & bé..." />
                </div>
                <div style="grid-column:1/-1;">
                  <label>Tình trạng mẹ (ghi chi tiết)</label>
                  <textarea id="scheduleMotherCondition" rows="3" placeholder="Tuần thai, triệu chứng, nhu cầu chăm sóc..."></textarea>
                </div>
                <div>
                  <label>Tình trạng bé</label>
                  <input id="scheduleBabyCondition" placeholder="Sức khỏe, ghi chú..." />
                </div>
                <div>
                  <label>Tư vấn viên (TV)</label>
                  <select id="scheduleConsultant"></select>
                </div>
                <div>
                  <label>Điều dưỡng (ĐD)</label>
                  <select id="scheduleNurse"></select>
                </div>
                <div>
                  <label>SALE</label>
                  <select id="scheduleSale"></select>
                </div>
                <div>
                  <label>Giá trải nghiệm (VNĐ)</label>
                  <input id="scheduleExpPrice" type="number" min="0" step="1000" placeholder="249000" />
                </div>
                <div>
                  <label>Thời gian buổi TN</label>
                  <input id="scheduleSessionDuration" placeholder="90p, 45p, 120p..." />
                </div>
                <div>
                  <label>Nguồn khách</label>
                  <input id="scheduleSource" placeholder="Facebook, Zalo, Giới thiệu..." />
                </div>
                <div>
                  <label>Kết quả HĐ – số tiền ký (VNĐ)</label>
                  <input id="scheduleContractAmount" type="number" min="0" step="1000" placeholder="0 nếu chưa ký" />
                </div>
                <div style="grid-column:1/-1;">
                  <label>Ghi chú thêm</label>
                  <textarea id="scheduleNote" rows="2" placeholder="Ghi chú nội bộ..."></textarea>
                </div>
                <div style="grid-column:1/-1;">
                  <button class="btn secondary" id="saveScheduleBtn" type="button">Lưu lịch</button>
                </div>
              </div>
            </div>
          </div>
        </section>

        <section class="card section app-page hidden" data-page="care" id="customerCareSection">
          <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
            <div>
              <h3>Chăm sóc khách hàng sau chốt</h3>
              <p class="muted" style="margin:2px 0 0;">Tự động lấy khách đã chốt từ Lịch khách hàng để theo dõi liệu trình và tiến độ sử dụng dịch vụ.</p>
            </div>
            <div style="display:flex;gap:8px;flex-wrap:wrap;">
              <button class="btn secondary" id="exportCareExcelBtn" type="button">Xuất Excel (CSV)</button>
              <button class="btn warn" id="exportCarePdfBtn" type="button">Xuất PDF CSKH</button>
            </div>
          </div>
          <div class="customer-filter form-grid" style="margin-top:10px;">
            <div class="customer-search-field">
              <label>Tìm nhanh (Tên / SĐT / Dịch vụ / Ghi chú)</label>
              <input id="careSearch" placeholder="Nhập từ khóa..." />
            </div>
            <div>
              <label>Từ ngày chốt</label>
              <input id="careFilterStartDate" type="date" />
            </div>
            <div>
              <label>Đến ngày chốt</label>
              <input id="careFilterEndDate" type="date" />
            </div>
            <div>
              <label>Nhân sự phụ trách (TV/ĐD/Sale)</label>
              <select id="careFilterStaff"></select>
            </div>
            <div>
              <label>Trạng thái CSKH</label>
              <select id="careFilterStatus"></select>
            </div>
            <div>
              <label>Nguồn data</label>
              <select id="careFilterSource"></select>
            </div>
            <div>
              <label>Tiến độ sử dụng</label>
              <select id="careFilterProgress">
                <option value="all">Tất cả tiến độ</option>
                <option value="not_started">Chưa sử dụng buổi nào</option>
                <option value="in_progress">Đang sử dụng</option>
                <option value="completed">Đã dùng hết liệu trình</option>
              </select>
            </div>
            <button class="btn secondary" id="applyCareFilterBtn" type="button">Áp dụng bộ lọc</button>
            <button class="btn secondary" id="resetCareFilterBtn" type="button">Đặt lại bộ lọc</button>
            <div class="muted" id="careFilterSummary" style="display:flex;align-items:center;font-size:0.83rem;">--</div>
          </div>

          <div class="tables" style="margin-top:10px;">
            <table>
              <thead>
                <tr>
                  <th>Ngày chốt</th>
                  <th>Khách hàng</th>
                  <th>SĐT</th>
                  <th>Dịch vụ</th>
                  <th>Tư vấn</th>
                  <th>Điều dưỡng</th>
                  <th>Sale</th>
                  <th>Nguồn</th>
                  <th>Giá trị HĐ</th>
                  <th>Liệu trình đăng ký</th>
                  <th>Đã sử dụng</th>
                  <th>Còn lại</th>
                  <th>Trạng thái CSKH</th>
                  <th>Hẹn chăm tiếp</th>
                  <th>Ghi chú CSKH</th>
                  <th>Thao tác</th>
                </tr>
              </thead>
              <tbody id="careBody"></tbody>
            </table>
          </div>
        </section>

        <section class="card section app-page hidden" data-page="accounting" id="accountingSection">
          <div id="accountingFolderView">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
              <div>
                <h3>Thư mục Kế toán</h3>
                <p class="muted" style="margin:2px 0 0;">Chọn nhóm nghiệp vụ kế toán để tiếp tục thao tác và quản lý dữ liệu.</p>
              </div>
            </div>
            <div class="reports-folder-grid" style="margin-top:12px;">
              <button class="reports-folder-card" type="button" data-accounting-folder="cashflow">
              <strong>Thu chi</strong>
              <span>Theo dõi phiếu thu, phiếu chi và đối soát dòng tiền hàng ngày.</span>
              </button>
              <button class="reports-folder-card" type="button" data-accounting-folder="attendance-office">
              <strong>Tính công văn phòng</strong>
              <span>Quản lý công làm việc, ca trực, ngày nghỉ cho nhân sự văn phòng.</span>
              </button>
              <button class="reports-folder-card" type="button" data-accounting-folder="attendance-service">
              <strong>Tính lương điều dưỡng</strong>
              <span>Quản lý công làm việc, ca trực, ngày nghỉ cho nhân sự dịch vụ.</span>
              </button>
              <button class="reports-folder-card" type="button" data-accounting-folder="payroll">
              <strong>Tính lương</strong>
              <span>Tính lương, phụ cấp, thưởng phạt và bảng lương nhân sự.</span>
              </button>
              <button class="reports-folder-card" type="button" data-accounting-folder="finance-report">
              <strong>Báo cáo tài chính</strong>
              <span>Tổng hợp doanh thu, chi phí, công nợ và báo cáo tài chính định kỳ.</span>
              </button>
            </div>
          </div>

          <div id="accountingDetailView" class="hidden">
            <div class="reports-detail-head">
              <div>
                <h3 id="accountingDetailTitle">Kế toán</h3>
                <p class="muted" id="accountingDetailDesc" style="margin:3px 0 0;">Nội dung mẫu theo từng nhóm nghiệp vụ.</p>
              </div>
              <div class="reports-detail-actions">
                <button class="btn secondary" id="accountingBackBtn" type="button">← Quay về thư mục</button>
              </div>
            </div>
            <div id="accountingDetailContent"></div>
          </div>
        </section>

        <section class="card section app-page hidden" data-page="access">
          <h3>Nguồn dữ liệu</h3>
          <div id="sourceSection">
            <div class="form-grid">
              <div>
                <label>Data Source</label>
                <select id="sourceType">
                  <option value="local">Local Browser Storage</option>
                  <option value="api">Backend REST API</option>
                  <option value="sheet">Google Sheets (CSV export)</option>
                </select>
              </div>
              <div>
                <label>Endpoint URL</label>
                <input id="sourceUrl" placeholder="https://api.example.com/reports hoặc https://docs.google.com/spreadsheets/.../pub?output=csv" />
              </div>
              <div style="display:flex;gap:8px;">
                <button class="btn secondary" id="syncBtn">Đồng bộ dữ liệu</button>
                <button class="btn secondary" id="testConnectionBtn">Kiểm tra kết nối</button>
              </div>
              <div class="muted" style="font-size:0.8rem;">API kỳ vọng: GET trả mảng reports JSON. Google Sheets: Publish as CSV và dùng link export.</div>
            </div>
          </div>

          <div id="telegramSourceSection" style="margin-top:18px;">
            <h3 style="margin:0 0 4px;">Đồng bộ từ Telegram</h3>
            <p class="muted" style="margin:0 0 10px;font-size:0.82rem;">Điều dưỡng gửi báo cáo ca vào nhóm Telegram theo mẫu quy định. Hệ thống đọc và nhập tự động vào danh sách lịch.</p>
            <div class="alert" style="margin-bottom:10px;font-size:0.82rem;">⚠ <strong>Lưu ý bảo mật:</strong> Bot Token được lưu trong trình duyệt. Không dùng token bot quan trọng khác. Chỉ Admin mới thấy mục này.</div>
            <div class="form-grid" style="grid-template-columns:1fr 1fr 1fr auto auto;align-items:end;">
              <div>
                <label>Bot Token (từ @BotFather)</label>
                <input id="telegramBotToken" type="password" placeholder="123456789:AAxxxxxxxxxxxxxx" />
              </div>
              <div>
                <label>Chat ID (ID nhóm/kênh)</label>
                <input id="telegramChatId" placeholder="-1001234567890" />
              </div>
              <div>
                <label>Webhook Public URL (https)</label>
                <input id="telegramWebhookBaseUrl" placeholder="https://your-public-domain.com" />
              </div>
              <button class="btn secondary" type="button" id="syncTelegramBtn">Đọc tin báo cáo</button>
              <button class="btn secondary" type="button" id="testTelegramBtn">Kết nối realtime</button>
            </div>
            <div class="muted" style="margin-top:6px;font-size:0.8rem;" id="telegramSyncStatus">Chưa đồng bộ lần nào.</div>
            <details style="margin-top:10px;">
              <summary style="cursor:pointer;font-weight:600;font-size:0.85rem;">📋 Hướng dẫn tạo bot và mẫu tin báo cáo</summary>
              <div class="muted" style="margin-top:8px;font-size:0.82rem;line-height:1.7;">
                <strong>Bước 1:</strong> Nhắn @BotFather trên Telegram → /newbot → lấy Token.<br />
                <strong>Bước 2:</strong> Thêm bot vào nhóm chat của điều dưỡng, cấp quyền đọc tin nhắn.<br />
                <strong>Bước 3:</strong> Lấy Chat ID: gọi <code>https://api.telegram.org/bot&lt;TOKEN&gt;/getUpdates</code> sau khi nhóm đã nhắn 1 tin. Có thể nhập nhiều Chat ID, ngăn cách bằng dấu phẩy.<br />
                <strong>Bước 4:</strong> Nhập Webhook Public URL (địa chỉ HTTPS Telegram truy cập được, ví dụ domain server hoặc tunnel URL).<br />
                <strong>Bước 5:</strong> Bấm <strong>Kết nối realtime</strong> để đăng ký webhook.<br />
                <strong>Bước 6:</strong> Điều dưỡng gửi tin theo mẫu dưới đây, hệ thống tự động nhận và cập nhật định kỳ.<br /><br />
                <strong>Mẫu hashtag phân luồng:</strong> <code>#dieuduong</code>, <code>#mkt</code>, <code>#tuvan</code>, <code>#telesale</code> (mỗi tin chỉ nên có 1 hashtag chính).<br />
                <strong>Mẫu tin điều dưỡng (gửi vào nhóm):</strong><br />
                <code style="display:block;background:#f8fafc;border:1px solid #dde5f0;border-radius:6px;padding:8px 10px;margin-top:4px;white-space:pre-wrap;">#baocao
ten: Nguyễn Thị Yến
ngay: 15/04/2026
gio: 09:00
khach: Trần Thị Anh
dichvu: Chăm sóc mẹ bầu
thoiluong: 90
khoangcach: 12
trangthai: hoàn tất</code>
                <strong>Mẫu tin marketing:</strong><br />
                <code style="display:block;background:#f8fafc;border:1px solid #dde5f0;border-radius:6px;padding:8px 10px;margin-top:4px;white-space:pre-wrap;">#mkt
ten: Lan Anh
ngay: 23/04/2026
khach: Lead Facebook
sdt: 0912345678
ngansach: 1500000
mess: 42
hopdong: 5000000</code>
                <br />Các trường bắt buộc: <strong>ten</strong> (điều dưỡng), <strong>ngay</strong>, <strong>khach</strong>.<br />
                Trường tuỳ chọn: gio, dichvu, thoiluong (phút), khoangcach (km), trangthai.<br />
                Mỗi báo cáo là 1 tin nhắn riêng và phải có hashtag phân luồng.
              </div>
            </details>
          </div>

          <div id="exportSection" style="margin-top:14px;">
            <h3>Xuất báo cáo</h3>
            <button class="btn warn" id="pdfBtn">Xuất PDF Dashboard</button>
          </div>
        </section>

        <section class="card section app-page hidden" data-page="activity" id="activitySection">
          <h3>Lịch sử hoạt động</h3>
          <p class="muted" id="activityHint">Chỉ Admin được xem lịch sử chỉnh sửa và thao tác trên hệ thống.</p>
          <div class="form-grid activity-controls" style="margin-bottom:10px;grid-template-columns:minmax(180px,240px) minmax(180px,240px) auto auto auto;align-items:end;">
            <div>
              <label>Từ ngày</label>
              <input id="activityFilterStartDate" type="date" value="${today}" />
            </div>
            <div>
              <label>Đến ngày</label>
              <input id="activityFilterEndDate" type="date" value="${today}" />
            </div>
            <button class="btn secondary" id="applyActivityFilterBtn" type="button">Áp dụng</button>
            <button class="btn secondary" id="resetActivityFilterBtn" type="button">Hôm nay</button>
            <div class="muted" id="activityPageInfo" style="font-size:0.82rem;display:flex;align-items:center;">Trang 1/1</div>
          </div>
          <div class="tables">
            <table>
              <thead>
                <tr><th>Thời gian</th><th>Người thực hiện</th><th>Phân hệ</th><th>Hành động</th><th>Chi tiết</th><th>Khôi phục</th></tr>
              </thead>
              <tbody id="activityBody"></tbody>
            </table>
          </div>
          <div class="login-actions" style="margin-top:10px;max-width:260px;">
            <button class="btn secondary" id="activityPrevPageBtn" type="button">← Trang trước</button>
            <button class="btn secondary" id="activityNextPageBtn" type="button">Trang sau →</button>
          </div>
        </section>
      </main>
    </div>
  </div>
  <div id="toastContainer"></div>
`;

const els = {
  dashboardRoot: document.querySelector("#dashboardRoot"),
  brandBlock: document.querySelector("#brandBlock"),
  brandTitleLogo: document.querySelector("#brandTitleLogo"),
  brandTitleText: document.querySelector("#brandTitleText"),
  headerActions: document.querySelector(".header-actions"),
  logoUpload: document.querySelector("#logoUpload"),
  menuToggle: document.querySelector("#menuToggle"),
  backBtn: document.querySelector("#backBtn"),
  menuOverlay: document.querySelector("#menuOverlay"),
  appMenu: document.querySelector("#appMenu"),
  menuLogoutBtn: document.querySelector("#menuLogoutBtn"),
  menuItems: Array.from(document.querySelectorAll(".menu-item")),
  appPages: Array.from(document.querySelectorAll(".app-page")),
  activitySection: document.querySelector("#activitySection"),
  activityBody: document.querySelector("#activityBody"),
  activityHint: document.querySelector("#activityHint"),
  activityFilterStartDate: document.querySelector("#activityFilterStartDate"),
  activityFilterEndDate: document.querySelector("#activityFilterEndDate"),
  applyActivityFilterBtn: document.querySelector("#applyActivityFilterBtn"),
  resetActivityFilterBtn: document.querySelector("#resetActivityFilterBtn"),
  activityPageInfo: document.querySelector("#activityPageInfo"),
  activityPrevPageBtn: document.querySelector("#activityPrevPageBtn"),
  activityNextPageBtn: document.querySelector("#activityNextPageBtn"),
  lastUpdated: document.querySelector("#lastUpdated"),
  loginUsername: document.querySelector("#loginUsername"),
  loginPassword: document.querySelector("#loginPassword"),
  toggleLoginPasswordBtn: document.querySelector("#toggleLoginPasswordBtn"),
  loginRemember: document.querySelector("#loginRemember"),
  loginBtn: document.querySelector("#loginBtn"),
  logoutBtn: document.querySelector("#logoutBtn"),
  authMessage: document.querySelector("#authMessage"),
  postLoginOnly: Array.from(document.querySelectorAll(".post-login-only")),
  mainContent: document.querySelector("#mainContent"),
  newsSection: document.querySelector("#newsSection"),
  newsComposeInputBtn: document.querySelector("#newsComposeInputBtn"),
  newsComposerEditor: document.querySelector("#newsComposerEditor"),
  newsComposerText: document.querySelector("#newsComposerText"),
  newsComposerAttachmentInput: document.querySelector("#newsComposerAttachmentInput"),
  newsComposerAttachmentList: document.querySelector("#newsComposerAttachmentList"),
  newsComposerDept: document.querySelector("#newsComposerDept"),
  newsComposerEventDate: document.querySelector("#newsComposerEventDate"),
  newsComposerImportant: document.querySelector("#newsComposerImportant"),
  newsSubmitPostBtn: document.querySelector("#newsSubmitPostBtn"),
  newsPostList: document.querySelector("#newsPostList"),
  newsEventFeedList: document.querySelector("#newsEventFeedList"),
  newsPinnedList: document.querySelector("#newsPinnedList"),
  newsEventsList: document.querySelector("#newsEventsList"),
  newsTrendList: document.querySelector("#newsTrendList"),
  newsEventModal: document.querySelector("#newsEventModal"),
  newsEventModalBackdrop: document.querySelector("#newsEventModalBackdrop"),
  closeNewsEventModalBtn: document.querySelector("#closeNewsEventModalBtn"),
  saveNewsEventBtn: document.querySelector("#saveNewsEventBtn"),
  newsEventTitle: document.querySelector("#newsEventTitle"),
  newsEventDate: document.querySelector("#newsEventDate"),
  newsEventTime: document.querySelector("#newsEventTime"),
  newsEventLocation: document.querySelector("#newsEventLocation"),
  newsEventDescription: document.querySelector("#newsEventDescription"),
  newsEventAttachmentInput: document.querySelector("#newsEventAttachmentInput"),
  newsEventAttachmentList: document.querySelector("#newsEventAttachmentList"),
  newsEventModalStatus: document.querySelector("#newsEventModalStatus"),
  newsVoteDetailModal: document.querySelector("#newsVoteDetailModal"),
  newsVoteDetailModalBackdrop: document.querySelector("#newsVoteDetailModalBackdrop"),
  closeNewsVoteDetailBtn: document.querySelector("#closeNewsVoteDetailBtn"),
  newsVoteDetailTitle: document.querySelector("#newsVoteDetailTitle"),
  newsVoteDetailList: document.querySelector("#newsVoteDetailList"),
  newsContentDetailModal: document.querySelector("#newsContentDetailModal"),
  newsContentDetailBackdrop: document.querySelector("#newsContentDetailBackdrop"),
  closeNewsContentDetailBtn: document.querySelector("#closeNewsContentDetailBtn"),
  newsContentDetailTitle: document.querySelector("#newsContentDetailTitle"),
  newsContentDetailMeta: document.querySelector("#newsContentDetailMeta"),
  newsContentDetailText: document.querySelector("#newsContentDetailText"),
  newsContentDetailImagePreview: document.querySelector("#newsContentDetailImagePreview"),
  newsContentDetailAttachments: document.querySelector("#newsContentDetailAttachments"),
  customerSection: document.querySelector("#customerSection"),
  adminUserSection: document.querySelector("#adminUserSection"),
  openUserModalBtn: document.querySelector("#openUserModalBtn"),
  userModal: document.querySelector("#userModal"),
  userModalBackdrop: document.querySelector("#userModalBackdrop"),
  closeUserModalBtn: document.querySelector("#closeUserModalBtn"),
  userModalTitle: document.querySelector("#userModalTitle"),
  hrQuickSearch: document.querySelector("#hrQuickSearch"),
  hrFilterRole: document.querySelector("#hrFilterRole"),
  hrFilterDept: document.querySelector("#hrFilterDept"),
  hrFilterStatus: document.querySelector("#hrFilterStatus"),
  userCode: document.querySelector("#userCode"),
  userFullName: document.querySelector("#userFullName"),
  userUsername: document.querySelector("#userUsername"),
  userPassword: document.querySelector("#userPassword"),
  userRoleKey: document.querySelector("#userRoleKey"),
  userDepartment: document.querySelector("#userDepartment"),
  userPhone: document.querySelector("#userPhone"),
  userEmail: document.querySelector("#userEmail"),
  userAddress: document.querySelector("#userAddress"),
  userBankAccount: document.querySelector("#userBankAccount"),
  saveUserBtn: document.querySelector("#saveUserBtn"),
  userManageStatus: document.querySelector("#userManageStatus"),
  userBody: document.querySelector("#userBody"),
  userModalStatus: document.querySelector("#userModalStatus"),
  hrProfileModal: document.querySelector("#hrProfileModal"),
  hrProfileModalBackdrop: document.querySelector("#hrProfileModalBackdrop"),
  closeHrProfileBtn: document.querySelector("#closeHrProfileBtn"),
  openPermissionModalBtn: document.querySelector("#openPermissionModalBtn"),
  permissionModal: document.querySelector("#permissionModal"),
  permissionModalBackdrop: document.querySelector("#permissionModalBackdrop"),
  closePermissionModalBtn: document.querySelector("#closePermissionModalBtn"),
  permissionRoleSelect: document.querySelector("#permissionRoleSelect"),
  permissionFeatureGrid: document.querySelector("#permissionFeatureGrid"),
  permissionPageGrid: document.querySelector("#permissionPageGrid"),
  permissionModalStatus: document.querySelector("#permissionModalStatus"),
  resetPermissionBtn: document.querySelector("#resetPermissionBtn"),
  savePermissionBtn: document.querySelector("#savePermissionBtn"),
  hrProfileTitle: document.querySelector("#hrProfileTitle"),
  hrProfileSubtitle: document.querySelector("#hrProfileSubtitle"),
  hrProfileInfo: document.querySelector("#hrProfileInfo"),
  hrProfileFileList: document.querySelector("#hrProfileFileList"),
  hrFileUpload: document.querySelector("#hrFileUpload"),
  hrProfileStatus: document.querySelector("#hrProfileStatus"),
  customerName: document.querySelector("#customerName"),
  openCustomerModalBtn: document.querySelector("#openCustomerModalBtn"),
  customerModal: document.querySelector("#customerModal"),
  customerModalBackdrop: document.querySelector("#customerModalBackdrop"),
  closeCustomerModalBtn: document.querySelector("#closeCustomerModalBtn"),
  customerModalTitle: document.querySelector("#customerModalTitle"),
  customerContactPerson: document.querySelector("#customerContactPerson"),
  customerPhone: document.querySelector("#customerPhone"),
  customerEmail: document.querySelector("#customerEmail"),
  customerAddress: document.querySelector("#customerAddress"),
  customerAddrProvince: document.querySelector("#customerAddrProvince"),
  customerAddrDistrict: document.querySelector("#customerAddrDistrict"),
  customerAddrWard: document.querySelector("#customerAddrWard"),
  customerTier: document.querySelector("#customerTier"),
  customerStatus: document.querySelector("#customerStatus"),
  customerOwner: document.querySelector("#customerOwner"),
  customerSource: document.querySelector("#customerSource"),
  customerDemand: document.querySelector("#customerDemand"),
  customerNote: document.querySelector("#customerNote"),
  saveCustomerBtn: document.querySelector("#saveCustomerBtn"),
  customerStatusMessage: document.querySelector("#customerStatusMessage"),
  customerQuickSearch: document.querySelector("#customerQuickSearch"),
  customerDateRangeField: document.querySelector("#customerDateRangeField"),
  customerDateRangeTrigger: document.querySelector("#customerDateRangeTrigger"),
  customerDateRangePopover: document.querySelector("#customerDateRangePopover"),
  customerDateRangePresets: document.querySelector("#customerDateRangePresets"),
  customerFilterStartDate: document.querySelector("#customerFilterStartDate"),
  customerFilterEndDate: document.querySelector("#customerFilterEndDate"),
  applyCustomerDateRangeBtn: document.querySelector("#applyCustomerDateRangeBtn"),
  clearCustomerDateRangeBtn: document.querySelector("#clearCustomerDateRangeBtn"),
  customerFilterOwner: document.querySelector("#customerFilterOwner"),
  customerFilterStatus: document.querySelector("#customerFilterStatus"),
  customerFilterSource: document.querySelector("#customerFilterSource"),
  applyCustomerFilterBtn: document.querySelector("#applyCustomerFilterBtn"),
  resetCustomerFilterBtn: document.querySelector("#resetCustomerFilterBtn"),
  exportCustomerExcelBtn: document.querySelector("#exportCustomerExcelBtn"),
  exportCustomerPdfBtn: document.querySelector("#exportCustomerPdfBtn"),
  customerFilterSummary: document.querySelector("#customerFilterSummary"),
  customerBody: document.querySelector("#customerBody"),
  inventorySection: document.querySelector("#inventorySection"),
  openInventoryModalBtn: document.querySelector("#openInventoryModalBtn"),
  openInventoryHistoryBtn: document.querySelector("#openInventoryHistoryBtn"),
  inventoryHistoryModal: document.querySelector("#inventoryHistoryModal"),
  inventoryHistoryModalBackdrop: document.querySelector("#inventoryHistoryModalBackdrop"),
  closeInventoryHistoryBtn: document.querySelector("#closeInventoryHistoryBtn"),
  exportInventoryExcelBtn: document.querySelector("#exportInventoryExcelBtn"),
  exportInventoryPdfBtn: document.querySelector("#exportInventoryPdfBtn"),
  inventorySearch: document.querySelector("#inventorySearch"),
  inventoryFilterStatus: document.querySelector("#inventoryFilterStatus"),
  inventoryFilterStock: document.querySelector("#inventoryFilterStock"),
  inventoryFilterSummary: document.querySelector("#inventoryFilterSummary"),
  inventoryTotalValue: document.querySelector("#inventoryTotalValue"),
  inventoryStatusMessage: document.querySelector("#inventoryStatusMessage"),
  inventoryStatsStart: document.querySelector("#inventoryStatsStart"),
  inventoryStatsEnd: document.querySelector("#inventoryStatsEnd"),
  applyInventoryStatsBtn: document.querySelector("#applyInventoryStatsBtn"),
  resetInventoryStatsBtn: document.querySelector("#resetInventoryStatsBtn"),
  inventoryStatsRangeText: document.querySelector("#inventoryStatsRangeText"),
  inventoryInTotal: document.querySelector("#inventoryInTotal"),
  inventoryOutTotal: document.querySelector("#inventoryOutTotal"),
  inventoryTxnTotal: document.querySelector("#inventoryTxnTotal"),
  inventoryNetTotal: document.querySelector("#inventoryNetTotal"),
  inventoryBody: document.querySelector("#inventoryBody"),
  inventoryHistoryBody: document.querySelector("#inventoryHistoryBody"),
  inventoryModal: document.querySelector("#inventoryModal"),
  inventoryModalBackdrop: document.querySelector("#inventoryModalBackdrop"),
  closeInventoryModalBtn: document.querySelector("#closeInventoryModalBtn"),
  inventoryModalTitle: document.querySelector("#inventoryModalTitle"),
  inventorySupplier: document.querySelector("#inventorySupplier"),
  inventoryExpiryDate: document.querySelector("#inventoryExpiryDate"),
  inventoryProductCode: document.querySelector("#inventoryProductCode"),
  inventoryProductName: document.querySelector("#inventoryProductName"),
  inventoryPurchasePrice: document.querySelector("#inventoryPurchasePrice"),
  inventorySalePrice: document.querySelector("#inventorySalePrice"),
  inventoryQuantity: document.querySelector("#inventoryQuantity"),
  inventoryAlertThreshold: document.querySelector("#inventoryAlertThreshold"),
  inventoryProductStatus: document.querySelector("#inventoryProductStatus"),
  saveInventoryBtn: document.querySelector("#saveInventoryBtn"),
  inventoryModalStatus: document.querySelector("#inventoryModalStatus"),
  inventoryTxnModal: document.querySelector("#inventoryTxnModal"),
  inventoryTxnModalBackdrop: document.querySelector("#inventoryTxnModalBackdrop"),
  closeInventoryTxnModalBtn: document.querySelector("#closeInventoryTxnModalBtn"),
  inventoryTxnTitle: document.querySelector("#inventoryTxnTitle"),
  inventoryTxnProduct: document.querySelector("#inventoryTxnProduct"),
  inventoryTxnType: document.querySelector("#inventoryTxnType"),
  inventoryTxnQty: document.querySelector("#inventoryTxnQty"),
  inventoryTxnNote: document.querySelector("#inventoryTxnNote"),
  saveInventoryTxnBtn: document.querySelector("#saveInventoryTxnBtn"),
  inventoryTxnStatus: document.querySelector("#inventoryTxnStatus"),
  sourceType: document.querySelector("#sourceType"),
  sourceUrl: document.querySelector("#sourceUrl"),
  syncBtn: document.querySelector("#syncBtn"),
  testConnectionBtn: document.querySelector("#testConnectionBtn"),
  pdfBtn: document.querySelector("#pdfBtn"),
  telegramBotToken: document.querySelector("#telegramBotToken"),
  telegramChatId: document.querySelector("#telegramChatId"),
  telegramWebhookBaseUrl: document.querySelector("#telegramWebhookBaseUrl"),
  syncTelegramBtn: document.querySelector("#syncTelegramBtn"),
  testTelegramBtn: document.querySelector("#testTelegramBtn"),
  telegramSyncStatus: document.querySelector("#telegramSyncStatus"),
  workflowSection: document.querySelector("#workflowSection"),
  workflowSystemView: document.querySelector("#workflowSystemView"),
  workflowDetailView: document.querySelector("#workflowDetailView"),
  workflowDetailTitle: document.querySelector("#workflowDetailTitle"),
  workflowDetailSteps: document.querySelector("#workflowDetailSteps"),
  workflowBackBtn: document.querySelector("#workflowBackBtn"),
  workflowDetailButtons: Array.from(document.querySelectorAll("[data-workflow-target]")),
  policySection: document.querySelector("#policySection"),
  policyFolderRoot: document.querySelector("#policyFolderRoot"),
  reportsSection: document.querySelector("#reportsSection"),
  reportsFolderView: document.querySelector("#reportsFolderView"),
  reportsDetailView: document.querySelector("#reportsDetailView"),
  reportsDetailTitle: document.querySelector("#reportsDetailTitle"),
  reportsDetailDesc: document.querySelector("#reportsDetailDesc"),
  reportsBackBtn: document.querySelector("#reportsBackBtn"),
  reportsSyncMetricsBtn: document.querySelector("#reportsSyncMetricsBtn"),
  exportReportsExcelBtn: document.querySelector("#exportReportsExcelBtn"),
  exportReportsPdfBtn: document.querySelector("#exportReportsPdfBtn"),
  reportsStartDate: document.querySelector("#reportsStartDate"),
  reportsEndDate: document.querySelector("#reportsEndDate"),
  reportsMarketingFilterWrap: document.querySelector("#reportsMarketingFilterWrap"),
  reportsMarketingFilter: document.querySelector("#reportsMarketingFilter"),
  reportsConsultantFilterWrap: document.querySelector("#reportsConsultantFilterWrap"),
  reportsConsultantFilter: document.querySelector("#reportsConsultantFilter"),
  reportsTelesaleFilterWrap: document.querySelector("#reportsTelesaleFilterWrap"),
  reportsTelesaleFilter: document.querySelector("#reportsTelesaleFilter"),
  applyReportsFilterBtn: document.querySelector("#applyReportsFilterBtn"),
  resetReportsFilterBtn: document.querySelector("#resetReportsFilterBtn"),
  reportsSummary: document.querySelector("#reportsSummary"),
  reportsTable: document.querySelector("#reportsTable"),
  reportsPrimaryCol: document.querySelector("#reportsPrimaryCol"),
  reportsMetricCol1: document.querySelector("#reportsMetricCol1"),
  reportsMetricCol2: document.querySelector("#reportsMetricCol2"),
  reportsMetricCol3: document.querySelector("#reportsMetricCol3"),
  reportsMetricCol4: document.querySelector("#reportsMetricCol4"),
  reportsBody: document.querySelector("#reportsBody"),
  reportForm: document.querySelector("#reportForm"),
  reportDate: document.querySelector("#reportDate"),
  submitReportBtn: document.querySelector("#submitReportBtn"),
  submitStatus: document.querySelector("#submitStatus"),
  timePreset: document.querySelector("#timePreset"),
  filterStart: document.querySelector("#filterStart"),
  filterEnd: document.querySelector("#filterEnd"),
  applyFilterBtn: document.querySelector("#applyFilterBtn"),
  resetFilterBtn: document.querySelector("#resetFilterBtn"),
  filterSummary: document.querySelector("#filterSummary"),
  kpiReports: document.querySelector("#kpiReports"),
  kpiCompletion: document.querySelector("#kpiCompletion"),
  kpiQuality: document.querySelector("#kpiQuality"),
  kpiIssues: document.querySelector("#kpiIssues"),
  homeDailyRevenue: document.querySelector("#homeDailyRevenue"),
  homeDailyRevenueMeta: document.querySelector("#homeDailyRevenueMeta"),
  homeMonthlyRevenue: document.querySelector("#homeMonthlyRevenue"),
  homeMonthlyRevenueMeta: document.querySelector("#homeMonthlyRevenueMeta"),
  homeTotalRevenue: document.querySelector("#homeTotalRevenue"),
  homeTotalRevenueMeta: document.querySelector("#homeTotalRevenueMeta"),
  homeDeptBody: document.querySelector("#homeDeptBody"),
  homeRevenueChart: document.querySelector("#homeRevenueChart"),
  metricsStartDate: document.querySelector("#metricsStartDate"),
  metricsEndDate: document.querySelector("#metricsEndDate"),
  metricsDepartmentFilter: document.querySelector("#metricsDepartmentFilter"),
  applyMetricsFilterBtn: document.querySelector("#applyMetricsFilterBtn"),
  resetMetricsFilterBtn: document.querySelector("#resetMetricsFilterBtn"),
  metricsFilterSummary: document.querySelector("#metricsFilterSummary"),
  mkInteractions: document.querySelector("#mkInteractions"),
  mkPhones: document.querySelector("#mkPhones"),
  mkLeadRate: document.querySelector("#mkLeadRate"),
  tsAppointments: document.querySelector("#tsAppointments"),
  tsBookingRate: document.querySelector("#tsBookingRate"),
  tsCancelRate: document.querySelector("#tsCancelRate"),
  tvRevenue: document.querySelector("#tvRevenue"),
  tvSignRate: document.querySelector("#tvSignRate"),
  tvAvgContract: document.querySelector("#tvAvgContract"),
  ddCaring: document.querySelector("#ddCaring"),
  ddServiceScore: document.querySelector("#ddServiceScore"),
  ddCompleteRate: document.querySelector("#ddCompleteRate"),
  metricsDeptKpiChart: document.querySelector("#metricsDeptKpiChart"),
  metricsDataFile: document.querySelector("#metricsDataFile"),
  metricsImportFileBtn: document.querySelector("#metricsImportFileBtn"),
  metricsSheetUrl: document.querySelector("#metricsSheetUrl"),
  metricsSyncSheetBtn: document.querySelector("#metricsSyncSheetBtn"),
  reportBody: document.querySelector("#reportBody"),
  chart: document.querySelector("#deptChart"),
  scheduleSection: document.querySelector("#scheduleSection"),
  openScheduleModalBtn: document.querySelector("#openScheduleModalBtn"),
  exportScheduleExcelBtn: document.querySelector("#exportScheduleExcelBtn"),
  exportSchedulePdfBtn: document.querySelector("#exportSchedulePdfBtn"),
  scheduleSearch: document.querySelector("#scheduleSearch"),
  scheduleFilterMonth: document.querySelector("#scheduleFilterMonth"),
  scheduleFilterStatus: document.querySelector("#scheduleFilterStatus"),
  scheduleFilterStaff: document.querySelector("#scheduleFilterStaff"),
  scheduleFilterSource: document.querySelector("#scheduleFilterSource"),
  applyScheduleFilterBtn: document.querySelector("#applyScheduleFilterBtn"),
  resetScheduleFilterBtn: document.querySelector("#resetScheduleFilterBtn"),
  scheduleFilterSummary: document.querySelector("#scheduleFilterSummary"),
  scheduleTableWrap: document.querySelector("#scheduleTableWrap"),
  scheduleTable: document.querySelector("#scheduleTable"),
  scheduleBody: document.querySelector("#scheduleBody"),
  scheduleScrollControls: document.querySelector("#scheduleScrollControls"),
  scheduleScrollLeftBtn: document.querySelector("#scheduleScrollLeftBtn"),
  scheduleScrollRightBtn: document.querySelector("#scheduleScrollRightBtn"),
  scheduleBottomScroller: document.querySelector("#scheduleBottomScroller"),
  scheduleBottomScrollerInner: document.querySelector("#scheduleBottomScrollerInner"),
  scheduleModal: document.querySelector("#scheduleModal"),
  scheduleModalBackdrop: document.querySelector("#scheduleModalBackdrop"),
  closeScheduleModalBtn: document.querySelector("#closeScheduleModalBtn"),
  scheduleModalTitle: document.querySelector("#scheduleModalTitle"),
  scheduleRegDate: document.querySelector("#scheduleRegDate"),
  scheduleTime: document.querySelector("#scheduleTime"),
  scheduleStatus: document.querySelector("#scheduleStatus"),
  scheduleName: document.querySelector("#scheduleName"),
  schedulePhone: document.querySelector("#schedulePhone"),
  scheduleMotherAge: document.querySelector("#scheduleMotherAge"),
  scheduleAddress: document.querySelector("#scheduleAddress"),
  scheduleBirthHistory: document.querySelector("#scheduleBirthHistory"),
  scheduleBabyBirthday: document.querySelector("#scheduleBabyBirthday"),
  scheduleStage: document.querySelector("#scheduleStage"),
  scheduleService: document.querySelector("#scheduleService"),
  scheduleMotherCondition: document.querySelector("#scheduleMotherCondition"),
  scheduleBabyCondition: document.querySelector("#scheduleBabyCondition"),
  scheduleConsultant: document.querySelector("#scheduleConsultant"),
  scheduleNurse: document.querySelector("#scheduleNurse"),
  scheduleSale: document.querySelector("#scheduleSale"),
  scheduleExpPrice: document.querySelector("#scheduleExpPrice"),
  scheduleSessionDuration: document.querySelector("#scheduleSessionDuration"),
  scheduleSource: document.querySelector("#scheduleSource"),
  scheduleContractAmount: document.querySelector("#scheduleContractAmount"),
  scheduleNote: document.querySelector("#scheduleNote"),
  saveScheduleBtn: document.querySelector("#saveScheduleBtn")
  ,
  customerCareSection: document.querySelector("#customerCareSection"),
  careSearch: document.querySelector("#careSearch"),
  careFilterStartDate: document.querySelector("#careFilterStartDate"),
  careFilterEndDate: document.querySelector("#careFilterEndDate"),
  careFilterStaff: document.querySelector("#careFilterStaff"),
  careFilterStatus: document.querySelector("#careFilterStatus"),
  careFilterSource: document.querySelector("#careFilterSource"),
  careFilterProgress: document.querySelector("#careFilterProgress"),
  applyCareFilterBtn: document.querySelector("#applyCareFilterBtn"),
  resetCareFilterBtn: document.querySelector("#resetCareFilterBtn"),
  exportCareExcelBtn: document.querySelector("#exportCareExcelBtn"),
  exportCarePdfBtn: document.querySelector("#exportCarePdfBtn"),
  careFilterSummary: document.querySelector("#careFilterSummary"),
  careBody: document.querySelector("#careBody"),
  accountingSection: document.querySelector("#accountingSection"),
  accountingFolderView: document.querySelector("#accountingFolderView"),
  accountingDetailView: document.querySelector("#accountingDetailView"),
  accountingDetailTitle: document.querySelector("#accountingDetailTitle"),
  accountingDetailDesc: document.querySelector("#accountingDetailDesc"),
  accountingDetailContent: document.querySelector("#accountingDetailContent"),
  accountingBackBtn: document.querySelector("#accountingBackBtn")
};

let chart;
let homeRevenueChart;
let metricsDeptKpiChart;
const CARE_STATUS_OPTIONS = ["Mới chốt", "Đang chăm sóc", "Cần gọi lại", "Đã hoàn tất", "Tạm dừng"];
let users = loadJSON(STORAGE.users, seedUsers);
let customers = loadJSON(STORAGE.customers, seedCustomers);
customers = customers.map((customer) => normalizeCustomer(customer));
let inventoryItems = loadJSON(STORAGE.inventoryItems, seedInventoryItems).map((item) => ({
  ...item,
  purchasePrice: Number(item.purchasePrice) || 0,
  salePrice: Number(item.salePrice) || 0,
  quantity: Math.max(0, Number(item.quantity) || 0),
  alertThreshold: Math.max(0, Number(item.alertThreshold) || 0),
  status: item.status === "inactive" ? "inactive" : "active",
  updatedAt: Number(item.updatedAt) || Date.now()
}));
let inventoryTransactions = loadJSON(STORAGE.inventoryTransactions, []).map((txn) => ({
  ...txn,
  quantity: Math.max(0, Number(txn.quantity) || 0),
  stockAfter: Math.max(0, Number(txn.stockAfter) || 0),
  createdAt: Number(txn.createdAt) || Date.now(),
  type: txn.type === "out" ? "out" : "in"
}));
let activityLogs = loadJSON(STORAGE.activities, []);
let recycleBin = loadJSON(STORAGE.recycleBin, []);
let hrFiles = loadJSON(STORAGE.hrFiles, {});
let authState = loadJSON(STORAGE.auth, { loggedIn: false, role: null, username: null, userId: null });
let reports = loadJSON(STORAGE.reports, seedReports);
let dataSourceConfig = normalizeDataSourceConfig(loadJSON(STORAGE.dataSource, { type: "local", url: "" }));
let usersSyncTimer = null;
let usersSyncPromise = null;
let usersSyncListenersBound = false;
const USERS_AUTO_SYNC_INTERVAL = 20000;
let cloudStorageStatus = { checkedAt: 0, durable: null, mode: "unknown", reason: "" };
let cloudStorageWarningShown = false;
let criticalStateSyncQueueTimer = null;
let criticalStatePullTimer = null;
let criticalStateSyncInFlight = false;
let criticalStateSyncListenersBound = false;
let isApplyingRemoteCriticalState = false;
const CRITICAL_STATE_AUTO_SYNC_INTERVAL = 30000;
let editingUserId = null;
let editingCustomerId = null;
let editingInventoryId = null;
let inventoryTxnItemId = null;
let inventoryStatsState = { start: "", end: "" };
let activePage = "home";
let pageHistory = [];
let filterState = { preset: "today", start: today, end: today };
let customerFilterState = normalizeCustomerFilterState(loadJSON(STORAGE.customerFilters, { start: "", end: "", owner: "all", status: "all", source: "all", keyword: "" }));
let rolePermissionsState = normalizeRolePermissions(loadJSON(STORAGE.rolePermissions, ROLES));
let permissionEditingRole = "staff";
let editingNewsPostId = null;
let editingNewsEventId = null;
let activeNewsDetail = null;
let pendingNewsPostAttachments = [];
let pendingNewsEventAttachments = [];
let activityViewState = {
  start: today,
  end: today,
  page: 1,
  pageSize: 30
};
let newsPosts = loadJSON(STORAGE.newsPosts, seedNewsPosts).map((item) => ({
  ...item,
  id: item.id || `np-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
  authorName: item.authorName || "Hệ thống",
  department: item.department || "Toàn hệ thống",
  content: String(item.content || "").trim(),
  tags: Array.isArray(item.tags) ? item.tags.filter(Boolean).slice(0, 4) : [],
  attachments: Array.isArray(item.attachments) ? item.attachments.map((file) => ({
    id: file.id || `nfa-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    name: String(file.name || "Tệp đính kèm"),
    type: String(file.type || "application/octet-stream"),
    data: String(file.data || ""),
    size: Math.max(0, Number(file.size) || 0)
  })).filter((file) => file.data) : [],
  views: Math.max(0, Number(item.views) || 0),
  comments: Math.max(0, Number(item.comments) || 0),
  tone: item.tone === "warn" || item.tone === "good" ? item.tone : "default",
  createdAt: Number(item.createdAt) || Date.now()
})).filter((item) => item.content);
let newsPinned = loadJSON(STORAGE.newsPinned, seedNewsPinned).map((item) => ({
  id: item.id || `pin-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
  text: String(item.text || "").trim(),
  createdAt: Number(item.createdAt) || Date.now()
})).filter((item) => item.text);
let newsEvents = loadJSON(STORAGE.newsEvents, seedNewsEvents).map((item) => ({
  id: item.id || `ne-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
  date: String(item.date || today),
  time: String(item.time || "09:00"),
  title: String(item.title || "").trim(),
  location: String(item.location || "").trim(),
  description: String(item.description || "").trim(),
  attachments: Array.isArray(item.attachments) ? item.attachments.map((file) => ({
    id: file.id || `nfa-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    name: String(file.name || "Tệp đính kèm"),
    type: String(file.type || "application/octet-stream"),
    data: String(file.data || ""),
    size: Math.max(0, Number(file.size) || 0)
  })).filter((file) => file.data) : [],
  votes: item.votes && typeof item.votes === "object" ? item.votes : {},
  createdAt: Number(item.createdAt) || Date.now()
})).filter((item) => item.title);
let activeWorkflowDetail = null;
let activePolicyParent = null;
let activePolicyChild = null;
let activeReportDepartment = null;
let activeAccountingFolder = null;
let reportFilterState = { start: `${today.slice(0, 7)}-01`, end: today };
let nurseReportSortState = { key: "total", direction: "desc" };
let marketingReportState = { marketing: "" };
let consultantReportState = { consultant: "", sortKey: "date", direction: "desc" };
let telesaleReportState = { sale: "", sortKey: "date", direction: "desc" };
let accountingCashflowEntries = loadJSON(STORAGE.accountingCashflow, seedAccountingCashflow).map((entry) => ({
  ...entry,
  amount: Math.max(0, Number(entry.amount) || 0),
  createdAt: Number(entry.createdAt) || Date.now(),
  type: entry.type === "expense" ? "expense" : "income",
  status: entry.status || "pending"
}));
let accountingCashflowFilterState = loadJSON(STORAGE.accountingCashflowFilters, {
  start: `${today.slice(0, 7)}-01`,
  end: today
});
let accountingAttendanceEntries = loadJSON(STORAGE.accountingAttendance, seedAccountingAttendance).map((entry) => ({
  ...entry,
  workHours: Math.max(0, Number(entry.workHours) || 0),
  overtimeHours: Math.max(0, Number(entry.overtimeHours) || 0),
  lateMinutes: Math.max(0, Number(entry.lateMinutes) || 0),
  date: String(entry.date || today).slice(0, 10)
}));
let accountingAttendanceSource = normalizeAccountingAttendanceSource(loadJSON(STORAGE.accountingAttendanceSource, {
  type: "sheet",
  url: "",
  vendor: "generic",
  autoSyncEnabled: false,
  autoSyncMinutes: 10,
  lastSyncedAt: 0,
  lastWarning: ""
}));
let accountingAttendanceFilterState = normalizeAccountingAttendanceFilterState(loadJSON(STORAGE.accountingAttendanceFilters, {
  start: `${today.slice(0, 7)}-01`,
  end: today
}));
let accountingServicePayrollFilterState = normalizeAccountingServicePayrollFilterState(loadJSON(STORAGE.accountingServicePayrollFilters, {
  start: `${today.slice(0, 7)}-01`,
  end: today
}));
let nurseReportOverrides = loadJSON(STORAGE.nurseReportOverrides, {});
let telegramSourceConfig = loadJSON(STORAGE.telegramSource, { token: "", chatId: "", webhookBaseUrl: "", lastUpdateId: 0, lastSyncedAt: 0 });
let attendanceAutoSyncTimer = null;
let attendanceSyncInProgress = false;

function getCurrentDataSourceConfig() {
  const type = els.sourceType ? els.sourceType.value : dataSourceConfig.type;
  const url = els.sourceUrl ? els.sourceUrl.value.trim() : dataSourceConfig.url;
  return normalizeDataSourceConfig({ type, url });
}

function saveDataSourceConfigFromInputs() {
  dataSourceConfig = getCurrentDataSourceConfig();
  saveJSON(STORAGE.dataSource, dataSourceConfig);
}

function applyDataSourceConfigToInputs() {
  if (els.sourceType) els.sourceType.value = dataSourceConfig.type;
  if (els.sourceUrl) els.sourceUrl.value = dataSourceConfig.url;
}

function isHttpUrl(value) {
  return /^https?:\/\//i.test(String(value || "").trim());
}

function deriveUsersEndpoint(rawUrl) {
  const normalized = String(rawUrl || "").trim().replace(/\/+$/g, "");
  if (!normalized) return "";
  if (/\/users?$/i.test(normalized)) return normalized;
  if (/\/reports?$/i.test(normalized)) return normalized.replace(/\/reports?$/i, "/users");
  return `${normalized}/users`;
}

function getSavedUsersSyncEndpoint() {
  return String(loadJSON(STORAGE.usersSyncEndpoint, "") || "").trim();
}

function saveUsersSyncEndpoint(endpointUrl) {
  saveJSON(STORAGE.usersSyncEndpoint, String(endpointUrl || "").trim());
}

function rememberUsersSyncEndpointFromSource() {
  const cfg = getCurrentDataSourceConfig();
  if (!isHttpUrl(cfg.url)) return;
  saveUsersSyncEndpoint(deriveUsersEndpoint(cfg.url));
}

async function loadRuntimeUsersSyncConfig() {
  try {
    const response = await fetch("./runtime-config.json", { method: "GET", cache: "no-store" });
    if (!response.ok) return;
    const config = await response.json();
    const endpointUrl = String(config?.usersSyncEndpoint || "").trim();
    if (!isHttpUrl(endpointUrl)) return;
    saveUsersSyncEndpoint(deriveUsersEndpoint(endpointUrl));
  } catch {
    // Optional config: ignore when runtime-config.json is missing or invalid.
  }
}

function getUsersSyncEndpoint() {
  const saved = getSavedUsersSyncEndpoint();
  if (isHttpUrl(saved)) return deriveUsersEndpoint(saved);

  const cfg = getCurrentDataSourceConfig();
  if (!isHttpUrl(cfg.url)) return "";
  return deriveUsersEndpoint(cfg.url);
}

const APP_STATE_SYNC_KEYS = new Set([
  STORAGE.customers,
  STORAGE.schedule,
  STORAGE.inventoryItems,
  STORAGE.inventoryTransactions,
  STORAGE.hrFiles,
  STORAGE.customerCareProgress,
  STORAGE.customerCareFilters,
  STORAGE.activities,
  STORAGE.recycleBin,
  STORAGE.rolePermissions,
  STORAGE.newsPosts,
  STORAGE.newsPinned,
  STORAGE.newsEvents,
  STORAGE.accountingCashflow,
  STORAGE.accountingCashflowFilters,
  STORAGE.accountingAttendance,
  STORAGE.accountingAttendanceSource,
  STORAGE.accountingAttendanceFilters,
  STORAGE.accountingServicePayrollFilters,
  STORAGE.nurseReportOverrides,
  STORAGE.telegramSource,
  STORAGE.dataSource,
  STORAGE.reports
]);

function deriveAppStateEndpoint(usersEndpoint) {
  const normalized = String(usersEndpoint || "").trim().replace(/\/+$/g, "");
  if (!normalized) return "";
  if (/\/app-state$/i.test(normalized)) return normalized;
  if (/\/users?$/i.test(normalized)) return normalized.replace(/\/users?$/i, "/app-state");
  return `${normalized}/app-state`;
}

function getAppStateSyncEndpoint() {
  const usersEndpoint = getUsersSyncEndpoint();
  if (!isHttpUrl(usersEndpoint)) return "";
  return deriveAppStateEndpoint(usersEndpoint);
}

function deriveStorageStatusEndpoint(usersEndpoint) {
  const normalized = String(usersEndpoint || "").trim().replace(/\/+$/g, "");
  if (!normalized) return "";
  if (/\/storage\/status$/i.test(normalized)) return normalized;
  if (/\/users?$/i.test(normalized)) return normalized.replace(/\/users?$/i, "/storage/status");
  return `${normalized}/storage/status`;
}

function isDurableStorageStatus(status = {}) {
  if (status?.durable === true) return true;
  if (String(status?.mode || "").toLowerCase() === "postgres") return true;
  const stateFile = String(status?.stateFile || "");
  return stateFile.startsWith("/var/data/");
}

function getUnsafeCloudStorageMessage(status = {}) {
  const mode = String(status?.mode || "json-file");
  const stateFile = String(status?.stateFile || "");
  if (mode === "postgres") return "Cloud backend chưa xác nhận durable=true.";
  if (stateFile) return `Cloud backend dang luu tam o ${stateFile}.`;
  return "Cloud backend chua o che do luu tru ben vung.";
}

function notifyUnsafeCloudStorage(reason, showToastMessage = false) {
  const message = reason || "Cloud backend chua o che do luu tru ben vung.";
  if (els.submitStatus) {
    els.submitStatus.textContent = `${message} Tam thoi chi giu local de tranh mat du lieu.`;
  }
  if (showToastMessage || !cloudStorageWarningShown) {
    showToast(`${message} Tam thoi chi giu local de tranh mat du lieu.`, "warning");
  }
  cloudStorageWarningShown = true;
}

async function ensureDurableCloudStorage(showWarning = false, forceRefresh = false) {
  const usersEndpoint = getUsersSyncEndpoint();
  if (!usersEndpoint) return true;

  const now = Date.now();
  if (!forceRefresh && cloudStorageStatus.checkedAt && now - cloudStorageStatus.checkedAt < 60000) {
    if (cloudStorageStatus.durable === false && showWarning) {
      notifyUnsafeCloudStorage(cloudStorageStatus.reason, true);
    }
    return cloudStorageStatus.durable !== false;
  }

  const endpointUrl = deriveStorageStatusEndpoint(usersEndpoint);
  if (!endpointUrl) return true;

  try {
    const response = await fetch(endpointUrl, { method: "GET", cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Không thể kiểm tra storage status (${response.status})`);
    }
    const status = await response.json();
    const durable = isDurableStorageStatus(status);
    cloudStorageStatus = {
      checkedAt: now,
      durable,
      mode: String(status?.mode || "unknown"),
      reason: durable ? "" : getUnsafeCloudStorageMessage(status)
    };
    if (!durable && showWarning) {
      notifyUnsafeCloudStorage(cloudStorageStatus.reason, true);
    }
    return durable;
  } catch (err) {
    cloudStorageStatus = {
      checkedAt: now,
      durable: false,
      mode: "unknown",
      reason: err.message || "Không thể kiểm tra cloud storage"
    };
    if (showWarning) {
      notifyUnsafeCloudStorage(cloudStorageStatus.reason, true);
    }
    return false;
  }
}

function countObjectKeys(value) {
  if (!value || typeof value !== "object") return 0;
  return Object.keys(value).length;
}

function hasLocalCriticalData() {
  return (
    (Array.isArray(customers) && customers.length > 0) ||
    (Array.isArray(schedules) && schedules.length > 0) ||
    (Array.isArray(inventoryItems) && inventoryItems.length > 0) ||
    (Array.isArray(inventoryTransactions) && inventoryTransactions.length > 0) ||
    countObjectKeys(hrFiles) > 0 ||
    countObjectKeys(customerCareProgress) > 0 ||
    (Array.isArray(activityLogs) && activityLogs.length > 0) ||
    (Array.isArray(recycleBin) && recycleBin.length > 0) ||
    (Array.isArray(newsPosts) && newsPosts.length > 0) ||
    (Array.isArray(newsPinned) && newsPinned.length > 0) ||
    (Array.isArray(newsEvents) && newsEvents.length > 0) ||
    (Array.isArray(accountingCashflowEntries) && accountingCashflowEntries.length > 0) ||
    (Array.isArray(accountingAttendanceEntries) && accountingAttendanceEntries.length > 0) ||
    countObjectKeys(nurseReportOverrides) > 0 ||
    (Array.isArray(reports) && reports.length > 0)
  );
}

function buildCriticalStatePayload() {
  return {
    schemaVersion: 2,
    customers: Array.isArray(customers) ? customers : [],
    schedules: Array.isArray(schedules) ? schedules : [],
    inventoryItems: Array.isArray(inventoryItems) ? inventoryItems : [],
    inventoryTransactions: Array.isArray(inventoryTransactions) ? inventoryTransactions : [],
    hrFiles: hrFiles && typeof hrFiles === "object" ? hrFiles : {},
    customerCareProgress: customerCareProgress && typeof customerCareProgress === "object" ? customerCareProgress : {},
    customerCareFilters: customerCareFilterState && typeof customerCareFilterState === "object" ? customerCareFilterState : {},
    activities: Array.isArray(activityLogs) ? activityLogs : [],
    recycleBin: Array.isArray(recycleBin) ? recycleBin : [],
    rolePermissions: rolePermissionsState && typeof rolePermissionsState === "object" ? rolePermissionsState : {},
    newsPosts: Array.isArray(newsPosts) ? newsPosts : [],
    newsPinned: Array.isArray(newsPinned) ? newsPinned : [],
    newsEvents: Array.isArray(newsEvents) ? newsEvents : [],
    accountingCashflow: Array.isArray(accountingCashflowEntries) ? accountingCashflowEntries : [],
    accountingCashflowFilters: accountingCashflowFilterState && typeof accountingCashflowFilterState === "object" ? accountingCashflowFilterState : {},
    accountingAttendance: Array.isArray(accountingAttendanceEntries) ? accountingAttendanceEntries : [],
    accountingAttendanceSource: accountingAttendanceSource && typeof accountingAttendanceSource === "object" ? accountingAttendanceSource : {},
    accountingAttendanceFilters: accountingAttendanceFilterState && typeof accountingAttendanceFilterState === "object" ? accountingAttendanceFilterState : {},
    accountingServicePayrollFilters: accountingServicePayrollFilterState && typeof accountingServicePayrollFilterState === "object" ? accountingServicePayrollFilterState : {},
    nurseReportOverrides: nurseReportOverrides && typeof nurseReportOverrides === "object" ? nurseReportOverrides : {},
    telegramSource: telegramSourceConfig && typeof telegramSourceConfig === "object" ? telegramSourceConfig : {},
    dataSourceConfig: dataSourceConfig && typeof dataSourceConfig === "object" ? dataSourceConfig : { type: "local", url: "" },
    reports: Array.isArray(reports) ? reports : [],
    updatedAt: Date.now()
  };
}

function normalizeRemoteCriticalState(raw = {}) {
  return {
    schemaVersion: Number(raw.schemaVersion) || 1,
    customers: Array.isArray(raw.customers) ? raw.customers : [],
    schedules: Array.isArray(raw.schedules) ? raw.schedules : [],
    inventoryItems: Array.isArray(raw.inventoryItems) ? raw.inventoryItems : [],
    inventoryTransactions: Array.isArray(raw.inventoryTransactions) ? raw.inventoryTransactions : [],
    hrFiles: raw.hrFiles && typeof raw.hrFiles === "object" ? raw.hrFiles : {},
    customerCareProgress: raw.customerCareProgress && typeof raw.customerCareProgress === "object" ? raw.customerCareProgress : {},
    customerCareFilters: raw.customerCareFilters && typeof raw.customerCareFilters === "object" ? raw.customerCareFilters : {},
    activities: Array.isArray(raw.activities) ? raw.activities : [],
    recycleBin: Array.isArray(raw.recycleBin) ? raw.recycleBin : [],
    rolePermissions: raw.rolePermissions && typeof raw.rolePermissions === "object" ? raw.rolePermissions : {},
    newsPosts: Array.isArray(raw.newsPosts) ? raw.newsPosts : [],
    newsPinned: Array.isArray(raw.newsPinned) ? raw.newsPinned : [],
    newsEvents: Array.isArray(raw.newsEvents) ? raw.newsEvents : [],
    accountingCashflow: Array.isArray(raw.accountingCashflow) ? raw.accountingCashflow : [],
    accountingCashflowFilters: raw.accountingCashflowFilters && typeof raw.accountingCashflowFilters === "object" ? raw.accountingCashflowFilters : {},
    accountingAttendance: Array.isArray(raw.accountingAttendance) ? raw.accountingAttendance : [],
    accountingAttendanceSource: raw.accountingAttendanceSource && typeof raw.accountingAttendanceSource === "object" ? raw.accountingAttendanceSource : {},
    accountingAttendanceFilters: raw.accountingAttendanceFilters && typeof raw.accountingAttendanceFilters === "object" ? raw.accountingAttendanceFilters : {},
    accountingServicePayrollFilters: raw.accountingServicePayrollFilters && typeof raw.accountingServicePayrollFilters === "object" ? raw.accountingServicePayrollFilters : {},
    nurseReportOverrides: raw.nurseReportOverrides && typeof raw.nurseReportOverrides === "object" ? raw.nurseReportOverrides : {},
    telegramSource: raw.telegramSource && typeof raw.telegramSource === "object" ? raw.telegramSource : {},
    dataSourceConfig: raw.dataSourceConfig && typeof raw.dataSourceConfig === "object" ? raw.dataSourceConfig : { type: "local", url: "" },
    reports: Array.isArray(raw.reports) ? raw.reports : [],
    updatedAt: Number(raw.updatedAt) || 0
  };
}

function getRemoteCriticalStateRowCount(remoteState) {
  return (
    (remoteState.customers?.length || 0) +
    (remoteState.schedules?.length || 0) +
    (remoteState.inventoryItems?.length || 0) +
    (remoteState.inventoryTransactions?.length || 0) +
    countObjectKeys(remoteState.hrFiles) +
    countObjectKeys(remoteState.customerCareProgress) +
    (remoteState.activities?.length || 0) +
    (remoteState.recycleBin?.length || 0) +
    (remoteState.newsPosts?.length || 0) +
    (remoteState.newsPinned?.length || 0) +
    (remoteState.newsEvents?.length || 0) +
    (remoteState.accountingCashflow?.length || 0) +
    (remoteState.accountingAttendance?.length || 0) +
    countObjectKeys(remoteState.nurseReportOverrides) +
    (remoteState.reports?.length || 0)
  );
}

function flushCriticalStateToRemoteWithBeacon() {
  const endpointUrl = getAppStateSyncEndpoint();
  if (!endpointUrl || typeof navigator.sendBeacon !== "function") return false;

  const pending = Boolean(loadJSON(STORAGE.criticalStatePendingSync, false));
  if (!pending) return false;

  try {
    const payload = JSON.stringify(buildCriticalStatePayload());
    const blob = new Blob([payload], { type: "application/json" });
    const sent = navigator.sendBeacon(endpointUrl, blob);
    if (sent) {
      localStorage.setItem(STORAGE.criticalStatePendingSync, "false");
    }
    return sent;
  } catch {
    return false;
  }
}

async function tryFlushCriticalStateWithKeepalive() {
  const endpointUrl = getAppStateSyncEndpoint();
  if (!endpointUrl) return false;
  const pending = Boolean(loadJSON(STORAGE.criticalStatePendingSync, false));
  if (!pending) return false;

  try {
    const response = await fetch(endpointUrl, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(buildCriticalStatePayload()),
      keepalive: true
    });
    if (response.ok) {
      localStorage.setItem(STORAGE.criticalStatePendingSync, "false");
      return true;
    }
  } catch {
    // Best-effort flush on background transitions.
  }
  return false;
}

async function fetchRemoteCriticalState(endpointUrl) {
  const response = await fetch(endpointUrl, { method: "GET" });
  if (!response.ok) throw new Error(`Không thể GET app-state (${response.status})`);
  const data = await response.json();
  return normalizeRemoteCriticalState(data);
}

async function pushRemoteCriticalState(endpointUrl) {
  const response = await fetch(endpointUrl, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(buildCriticalStatePayload())
  });
  if (!response.ok) throw new Error(`Không thể PUT app-state (${response.status})`);
}

async function syncCriticalStateToRemote(showToastOnSuccess = false) {
  if (criticalStateSyncInFlight) return false;
  const endpointUrl = getAppStateSyncEndpoint();
  if (!endpointUrl) return false;
  const storageReady = await ensureDurableCloudStorage(showToastOnSuccess);
  if (!storageReady) {
    localStorage.setItem(STORAGE.criticalStatePendingSync, "true");
    return false;
  }

  criticalStateSyncInFlight = true;
  try {
    await pushRemoteCriticalState(endpointUrl);
    localStorage.setItem(STORAGE.criticalStatePendingSync, "false");
    if (showToastOnSuccess) {
      showToast("Đã đồng bộ toàn bộ dữ liệu nghiệp vụ lên cloud.", "success");
    }
    return true;
  } catch (err) {
    localStorage.setItem(STORAGE.criticalStatePendingSync, "true");
    if (showToastOnSuccess) {
      showToast(`Đồng bộ cloud thất bại: ${err.message}`, "warning");
    }
    return false;
  } finally {
    criticalStateSyncInFlight = false;
  }
}

function queueCriticalStateSync(storageKey) {
  if (isApplyingRemoteCriticalState) return;
  if (!APP_STATE_SYNC_KEYS.has(storageKey)) return;
  localStorage.setItem(STORAGE.criticalStatePendingSync, "true");
  if (criticalStateSyncQueueTimer) clearTimeout(criticalStateSyncQueueTimer);
  criticalStateSyncQueueTimer = setTimeout(() => {
    syncCriticalStateToRemote(false).catch(() => {
      // Keep local data and retry on next cycle.
    });
  }, 1200);
}

async function syncCriticalStateFromRemote(showToastOnSuccess = false) {
  const endpointUrl = getAppStateSyncEndpoint();
  if (!endpointUrl) return false;
  const storageReady = await ensureDurableCloudStorage(showToastOnSuccess);
  if (!storageReady) return false;

  const hasPendingSync = Boolean(loadJSON(STORAGE.criticalStatePendingSync, false));
  if (hasPendingSync) {
    const pushed = await syncCriticalStateToRemote(showToastOnSuccess);
    if (pushed) return true;
    return false;
  }

  const remoteState = await fetchRemoteCriticalState(endpointUrl);
  const remoteRows = getRemoteCriticalStateRowCount(remoteState);
  if (remoteRows === 0 && hasLocalCriticalData()) {
    const pushed = await syncCriticalStateToRemote(showToastOnSuccess);
    if (!pushed) return false;
    return true;
  }

  isApplyingRemoteCriticalState = true;
  try {
    customers = remoteState.customers.map((customer) => normalizeCustomer(customer));
    schedules = remoteState.schedules;
    inventoryItems = remoteState.inventoryItems.map((item) => ({
      ...item,
      purchasePrice: Number(item.purchasePrice) || 0,
      salePrice: Number(item.salePrice) || 0,
      quantity: Math.max(0, Number(item.quantity) || 0),
      alertThreshold: Math.max(0, Number(item.alertThreshold) || 0),
      status: item.status === "inactive" ? "inactive" : "active",
      updatedAt: Number(item.updatedAt) || Date.now()
    }));
    inventoryTransactions = remoteState.inventoryTransactions.map((txn) => ({
      ...txn,
      quantity: Math.max(0, Number(txn.quantity) || 0),
      stockAfter: Math.max(0, Number(txn.stockAfter) || 0),
      createdAt: Number(txn.createdAt) || Date.now(),
      type: txn.type === "out" ? "out" : "in"
    }));
    hrFiles = remoteState.hrFiles;
    customerCareProgress = remoteState.customerCareProgress;
    customerCareFilterState = normalizeCustomerFilterState(remoteState.customerCareFilters);
    activityLogs = remoteState.activities;
    recycleBin = remoteState.recycleBin;
    rolePermissionsState = normalizeRolePermissions(remoteState.rolePermissions, ROLES);
    newsPosts = remoteState.newsPosts;
    newsPinned = remoteState.newsPinned;
    newsEvents = remoteState.newsEvents;
    accountingCashflowEntries = remoteState.accountingCashflow.map((entry) => ({
      ...entry,
      amount: Math.max(0, Number(entry.amount) || 0),
      createdAt: Number(entry.createdAt) || Date.now(),
      type: entry.type === "expense" ? "expense" : "income",
      status: entry.status || "pending"
    }));
    accountingCashflowFilterState = remoteState.accountingCashflowFilters;
    accountingAttendanceEntries = remoteState.accountingAttendance.map((entry) => ({
      ...entry,
      workHours: Math.max(0, Number(entry.workHours) || 0),
      overtimeHours: Math.max(0, Number(entry.overtimeHours) || 0),
      lateMinutes: Math.max(0, Number(entry.lateMinutes) || 0),
      date: String(entry.date || today).slice(0, 10)
    }));
    accountingAttendanceSource = normalizeAccountingAttendanceSource(remoteState.accountingAttendanceSource);
    accountingAttendanceFilterState = normalizeAccountingAttendanceFilterState(remoteState.accountingAttendanceFilters);
    accountingServicePayrollFilterState = normalizeAccountingServicePayrollFilterState(remoteState.accountingServicePayrollFilters);
    nurseReportOverrides = remoteState.nurseReportOverrides;
    telegramSourceConfig = remoteState.telegramSource;
    dataSourceConfig = normalizeDataSourceConfig(remoteState.dataSourceConfig);
    reports = remoteState.reports;

    saveJSON(STORAGE.customers, customers);
    saveJSON(STORAGE.schedule, schedules);
    saveJSON(STORAGE.inventoryItems, inventoryItems);
    saveJSON(STORAGE.inventoryTransactions, inventoryTransactions);
    saveJSON(STORAGE.hrFiles, hrFiles);
    saveJSON(STORAGE.customerCareProgress, customerCareProgress);
    saveJSON(STORAGE.customerCareFilters, customerCareFilterState);
    saveJSON(STORAGE.activities, activityLogs);
    saveJSON(STORAGE.recycleBin, recycleBin);
    saveJSON(STORAGE.rolePermissions, rolePermissionsState);
    saveJSON(STORAGE.newsPosts, newsPosts);
    saveJSON(STORAGE.newsPinned, newsPinned);
    saveJSON(STORAGE.newsEvents, newsEvents);
    saveJSON(STORAGE.accountingCashflow, accountingCashflowEntries);
    saveJSON(STORAGE.accountingCashflowFilters, accountingCashflowFilterState);
    saveJSON(STORAGE.accountingAttendance, accountingAttendanceEntries);
    saveJSON(STORAGE.accountingAttendanceSource, accountingAttendanceSource);
    saveJSON(STORAGE.accountingAttendanceFilters, accountingAttendanceFilterState);
    saveJSON(STORAGE.accountingServicePayrollFilters, accountingServicePayrollFilterState);
    saveJSON(STORAGE.nurseReportOverrides, nurseReportOverrides);
    saveJSON(STORAGE.telegramSource, telegramSourceConfig);
    saveJSON(STORAGE.dataSource, dataSourceConfig);
    saveJSON(STORAGE.reports, reports);
    localStorage.setItem(STORAGE.criticalStatePendingSync, "false");
  } finally {
    isApplyingRemoteCriticalState = false;
  }

  applyDataSourceConfigToInputs();
  rememberUsersSyncEndpointFromSource();

  if (showToastOnSuccess) {
    showToast("Đã tải toàn bộ dữ liệu nghiệp vụ từ cloud.", "success");
  }
  return true;
}

function startCriticalStateAutoSync() {
  if (criticalStatePullTimer) clearInterval(criticalStatePullTimer);
  criticalStatePullTimer = setInterval(() => {
    if (document.hidden) return;
    syncCriticalStateFromRemote(false).catch(() => {
      // Keep local fallback and retry.
    });
  }, CRITICAL_STATE_AUTO_SYNC_INTERVAL);

  if (criticalStateSyncListenersBound) return;
  criticalStateSyncListenersBound = true;

  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) {
      syncCriticalStateFromRemote(false).catch(() => {
        // Keep local fallback and retry.
      });
    }
  });

  window.addEventListener("focus", () => {
    syncCriticalStateFromRemote(false).catch(() => {
      // Keep local fallback and retry.
    });
  });

  window.addEventListener("online", () => {
    syncCriticalStateToRemote(false).catch(() => {
      // Retry on next cycle if online sync fails.
    });
  });

  window.addEventListener("pagehide", () => {
    if (!flushCriticalStateToRemoteWithBeacon()) {
      void tryFlushCriticalStateWithKeepalive();
    }
  });

  window.addEventListener("beforeunload", () => {
    flushCriticalStateToRemoteWithBeacon();
  });
}

function normalizeRemoteUser(user = {}) {
  return {
    id: String(user.id || `u-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`),
    userCode: String(user.userCode || "").trim().toUpperCase(),
    username: String(user.username || "").trim().toLowerCase(),
    password: String(user.password || ""),
    fullName: String(user.fullName || "").trim(),
    roleKey: String(user.roleKey || "staff"),
    department: String(user.department || "Ban điều hành"),
    phone: String(user.phone || ""),
    email: String(user.email || ""),
    address: String(user.address || ""),
    bankAccount: String(user.bankAccount || ""),
    status: user.status === "suspended" ? "suspended" : "active",
    createdAt: Number(user.createdAt) || Date.now()
  };
}

function sanitizeUsersForRemote(list) {
  return list.map((user) => normalizeRemoteUser(user));
}

function isDefaultUsersSnapshot(list = []) {
  if (!Array.isArray(list)) return false;
  if (list.length > 3) return false;
  const usernames = new Set(list.map((user) => String(user.username || "").toLowerCase()));
  return usernames.has("admin") && usernames.has("ceo") && usernames.has("head-tech");
}

function hasRicherLocalUsers(localList = [], remoteList = []) {
  if (!Array.isArray(localList) || !Array.isArray(remoteList)) return false;
  if (localList.length <= remoteList.length) return false;
  const remoteUsernames = new Set(remoteList.map((u) => String(u.username || "").toLowerCase()));
  return localList.some((u) => !remoteUsernames.has(String(u.username || "").toLowerCase()));
}

async function fetchRemoteUsers(endpointUrl) {
  const response = await fetch(endpointUrl, { method: "GET" });
  if (!response.ok) throw new Error(`Không thể GET users (${response.status})`);
  const data = await response.json();
  if (!Array.isArray(data)) throw new Error("Users remote phải là mảng JSON");
  return data.map((user) => normalizeRemoteUser(user)).filter((user) => user.username);
}

async function pushRemoteUsers(endpointUrl, usersList) {
  const payload = sanitizeUsersForRemote(usersList);
  const tryRequests = [
    { method: "PUT", body: JSON.stringify(payload) },
    { method: "POST", body: JSON.stringify(payload) },
    { method: "POST", body: JSON.stringify({ users: payload }) }
  ];

  let lastStatus = "N/A";
  for (const req of tryRequests) {
    const response = await fetch(endpointUrl, {
      method: req.method,
      headers: { "Content-Type": "application/json" },
      body: req.body
    });
    lastStatus = String(response.status);
    if (response.ok) return;
  }
  throw new Error(`Không thể đồng bộ users lên remote (HTTP ${lastStatus})`);
}

async function syncUsersFromRemote(showToastOnSuccess = false) {
  if (usersSyncPromise) return usersSyncPromise;

  usersSyncPromise = (async () => {
    const endpointUrl = getUsersSyncEndpoint();
    if (!endpointUrl) return false;
    const storageReady = await ensureDurableCloudStorage(showToastOnSuccess);
    if (!storageReady) return false;
    const remoteUsers = await fetchRemoteUsers(endpointUrl);
    if (!remoteUsers.length) return false;

    const hasPendingUsersSync = Boolean(loadJSON(STORAGE.usersPendingSync, false));
    const shouldRecoverFromReset =
      users.length &&
      hasRicherLocalUsers(users, remoteUsers) &&
      isDefaultUsersSnapshot(remoteUsers);

    if ((hasPendingUsersSync && users.length && remoteUsers.length < users.length) || shouldRecoverFromReset) {
      try {
        await pushRemoteUsers(endpointUrl, users);
        saveJSON(STORAGE.usersPendingSync, false);
        if (showToastOnSuccess) {
          showToast("Đã khôi phục users local lên cloud thành công.", "success");
        }
        return true;
      } catch (err) {
        showToast(`Cloud users chưa cập nhật, tạm giữ dữ liệu local để tránh mất dữ liệu: ${err.message}`, "warning");
        return false;
      }
    }

    users = remoteUsers;
    saveJSON(STORAGE.users, users);
    saveJSON(STORAGE.usersPendingSync, false);
    const forcedLogout = enforceActiveSessionAccess();
    if (forcedLogout) return true;
    if (showToastOnSuccess) {
      showToast(`Đã tải ${remoteUsers.length} tài khoản từ remote.`, "success");
    }
    return true;
  })();

  try {
    return await usersSyncPromise;
  } finally {
    usersSyncPromise = null;
  }
}

async function persistUsersToRemote(actionLabel = "") {
  saveJSON(STORAGE.users, users);
  saveJSON(STORAGE.usersPendingSync, true);
  const endpointUrl = getUsersSyncEndpoint();
  if (!endpointUrl) return true;
  const storageReady = await ensureDurableCloudStorage(true);
  if (!storageReady) {
    if (actionLabel) {
      logActivity("Nhân sự", "Cloud chưa bền vững", actionLabel);
    }
    return false;
  }
  try {
    await pushRemoteUsers(endpointUrl, users);
    const remoteUsers = await fetchRemoteUsers(endpointUrl);
    const localSignature = JSON.stringify(sanitizeUsersForRemote(users));
    const remoteSignature = JSON.stringify(sanitizeUsersForRemote(remoteUsers));
    if (localSignature !== remoteSignature) {
      throw new Error("Cloud chua xac nhan du lieu users moi nhat");
    }
    users = remoteUsers;
    saveJSON(STORAGE.users, users);
    saveJSON(STORAGE.usersPendingSync, false);
    return true;
  } catch (err) {
    showToast(`Đã lưu local nhưng chưa đồng bộ cloud: ${err.message}`, "warning");
    if (actionLabel) {
      logActivity("Nhân sự", "Cảnh báo đồng bộ cloud", `${actionLabel} | ${err.message}`);
    }
    return false;
  }
}

function syncUsersToRemoteInBackground(actionLabel = "") {
  const endpointUrl = getUsersSyncEndpoint();
  if (!endpointUrl) return;
  ensureDurableCloudStorage(false).then((storageReady) => {
    if (!storageReady) return;
    return pushRemoteUsers(endpointUrl, users);
  }).catch((err) => {
    showToast(`Chưa đồng bộ cloud users: ${err.message}`, "warning");
    if (actionLabel) {
      logActivity("Nhân sự", "Cảnh báo đồng bộ cloud", `${actionLabel} | ${err.message}`);
    }
  });
}

function startUsersAutoSync() {
  if (usersSyncTimer) clearInterval(usersSyncTimer);
  usersSyncTimer = setInterval(() => {
    if (document.hidden) return;
    syncUsersFromRemote(false).catch(() => {
      // Keep local fallback when endpoint is temporarily unavailable.
    });
  }, USERS_AUTO_SYNC_INTERVAL);

  if (usersSyncListenersBound) return;
  usersSyncListenersBound = true;

  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) {
      syncUsersFromRemote(false).catch(() => {
        // Keep local fallback when endpoint is temporarily unavailable.
      });
    }
  });

  window.addEventListener("focus", () => {
    syncUsersFromRemote(false).catch(() => {
      // Keep local fallback when endpoint is temporarily unavailable.
    });
  });
}
let attendanceAutoSyncSignature = "";
let telegramRealtimeSyncTimer = null;
let schedules = loadJSON(STORAGE.schedule, seedSchedule);
if (Array.isArray(schedules)) {
  const existingIds = new Set(schedules.map((item) => item.id));
  const missingSamples = seedSchedule.filter((item) => !existingIds.has(item.id));
  if (missingSamples.length) {
    schedules = [...schedules, ...missingSamples];
    saveJSON(STORAGE.schedule, schedules);
  }
}
let editingScheduleId = null;
let scheduleFilterState = { month: today.slice(0, 7), status: "", staff: "all", source: "", keyword: "" };
let metricsFilterState = { start: filterState.start, end: filterState.end, department: "all" };
let customerCareProgress = loadJSON(STORAGE.customerCareProgress, {});
let customerCareFilterState = normalizeCustomerCareFilterState(loadJSON(STORAGE.customerCareFilters, {
  start: "",
  end: "",
  staff: "all",
  status: "all",
  source: "all",
  progress: "all",
  keyword: ""
}));
enableHorizontalDragScroll(els.scheduleTableWrap);

let scheduleScrollSyncLock = false;
const SCHEDULE_SCROLL_STEP = 280;

function scrollScheduleHorizontally(delta) {
  if (!els.scheduleTableWrap) return;
  els.scheduleTableWrap.scrollBy({ left: delta, behavior: "smooth" });
}

function syncScheduleBottomScrollerWidth() {
  if (!els.scheduleTableWrap || !els.scheduleTable || !els.scheduleBottomScroller || !els.scheduleBottomScrollerInner) return;
  const tableWidth = els.scheduleTable.scrollWidth;
  const wrapWidth = els.scheduleTableWrap.clientWidth;
  els.scheduleBottomScrollerInner.style.width = `${tableWidth}px`;
  const shouldShow = tableWidth > wrapWidth + 2;
  els.scheduleScrollControls.classList.toggle("hidden", !shouldShow);
}

function initScheduleBottomScroller() {
  if (!els.scheduleTableWrap || !els.scheduleBottomScroller) return;

  els.scheduleTableWrap.addEventListener("scroll", () => {
    if (scheduleScrollSyncLock) return;
    scheduleScrollSyncLock = true;
    els.scheduleBottomScroller.scrollLeft = els.scheduleTableWrap.scrollLeft;
    scheduleScrollSyncLock = false;
  });

  els.scheduleBottomScroller.addEventListener("scroll", () => {
    if (scheduleScrollSyncLock) return;
    scheduleScrollSyncLock = true;
    els.scheduleTableWrap.scrollLeft = els.scheduleBottomScroller.scrollLeft;
    scheduleScrollSyncLock = false;
  });

  els.scheduleScrollLeftBtn.addEventListener("click", () => {
    scrollScheduleHorizontally(-SCHEDULE_SCROLL_STEP);
  });

  els.scheduleScrollRightBtn.addEventListener("click", () => {
    scrollScheduleHorizontally(SCHEDULE_SCROLL_STEP);
  });

  document.addEventListener("keydown", (event) => {
    const isScheduleVisible = activePage === "schedule" && els.scheduleSection && !els.scheduleSection.classList.contains("hidden");
    if (!isScheduleVisible) return;
    const target = event.target;
    if (target instanceof HTMLElement && target.closest("input,textarea,select")) return;

    if (event.key === "ArrowLeft" || event.key.toLowerCase() === "a") {
      event.preventDefault();
      scrollScheduleHorizontally(-SCHEDULE_SCROLL_STEP);
    } else if (event.key === "ArrowRight" || event.key.toLowerCase() === "d") {
      event.preventDefault();
      scrollScheduleHorizontally(SCHEDULE_SCROLL_STEP);
    }
  }, true);

  window.addEventListener("resize", syncScheduleBottomScrollerWidth);
  syncScheduleBottomScrollerWidth();
}

initScheduleBottomScroller();

const WORKFLOW_DEPARTMENT_DETAILS = {
  "hcns": {
    title: "Phong HCNS",
    steps: [
      "Xac dinh nhu cau nhan su theo tung bo phan va lap ke hoach tuyen dung.",
      "Dang tin, sang loc ho so, to chuc phong van va danh gia ung vien.",
      "Hoan thien hop dong, ho so lao dong va quy trinh tiep nhan nhan vien moi.",
      "Theo doi KPI, dao tao dinh ky va danh gia nang luc theo quy che cong ty.",
      "Quan ly che do luong thuong, phep, bao hiem va phuc loi noi bo.",
      "Cap nhat noi quy, van hoa doanh nghiep va xu ly van de nhan su phat sinh."
    ]
  },
  "marketing": {
    title: "Phong Marketing",
    steps: [
      "Nghien cuu thi truong, phan tich hanh vi khach hang va xu huong dich vu.",
      "Xay dung chien luoc va ke hoach truyen thong theo muc tieu tung giai doan.",
      "Trien khai noi dung tren cac kenh online/offline va theo doi hieu qua.",
      "Toi uu ngan sach quang cao va toc do chuyen doi lead.",
      "Tong hop feedback tu Kinh Doanh/CSKH de dieu chinh thong diep.",
      "Ban giao lead dat chuan vao CRM cho bo phan Kinh Doanh."
    ]
  },
  "kinh-doanh": {
    title: "Phong Kinh Doanh",
    steps: [
      "Tiep nhan lead tu Marketing, xep hang uu tien va phan bo cho tu van vien.",
      "Lien he xac nhan nhu cau, tu van giai phap va hen lich tuong tac.",
      "Trinh bay goi dich vu, xu ly phan doi va chot phuong an.",
      "Cap nhat ket qua len he thong va theo doi cac co hoi chua chot.",
      "Phan hoi du lieu thi truong ve Marketing de toi uu chien dich.",
      "Phoi hop Dieu Hanh Ca de ban giao thong tin dich vu da chot."
    ]
  },
  "tu-van": {
    title: "Bo phan Tu Van",
    steps: [
      "Nhan thong tin khach hang va lich hen tu Kinh Doanh/Dieu Hanh Ca.",
      "Khai thac van de, tu van lo trinh va de xuat phac do phu hop.",
      "Giai thich chi phi, quyen loi va cam ket chat luong truoc khi thuc hien.",
      "Ghi nhan thong tin quan trong va dong bo cho Dieu Duong va CSKH.",
      "Theo doi tinh trang khach sau tu van, xu ly cau hoi bo sung.",
      "Phoi hop chot lich va chuyen sang buoc van hanh ca."
    ]
  },
  "dieu-hanh-ca": {
    title: "Bo phan Dieu Hanh Ca",
    steps: [
      "Tong hop lich hen trong ngay va dieu pho nguon luc theo khung gio.",
      "Phan ca cho Dieu Duong/Tu Van theo nang luc va muc do uu tien.",
      "Kiem soat thong tin truoc ca, trong ca va sau ca de tranh sai lech.",
      "Xu ly dieu chinh lich dot xuat va tinh huong phat sinh tai cho.",
      "Ban giao thong tin ket qua cho CSKH va cac bo phan lien quan.",
      "Bao cao tong ket van hanh ca cho cap quan ly."
    ]
  },
  "dieu-duong": {
    title: "Bo phan Dieu Duong",
    steps: [
      "Nhan ca theo lenh Dieu Hanh, kiem tra thong tin va dung cu truoc xu ly.",
      "Thuc hien quy trinh chuyen mon dung tieu chuan ky thuat va an toan.",
      "Theo doi chi so, ghi nhan check-in/check-out va nhat ky ca.",
      "Bao cao bat thuong ngay trong ca cho Dieu Hanh va cap quan ly.",
      "Huong dan cham soc sau dich vu cho khach hang.",
      "Ban giao ket qua sau ca va cap nhat de CSKH tiep tuc theo doi."
    ]
  },
  "cskh": {
    title: "Bo phan Cham Soc Khach Hang",
    steps: [
      "Tiep nhan thong tin sau dich vu va lap lich cham soc dinh ky.",
      "Goi/xac nhan muc do hai long va thu thap phan hoi chat luong.",
      "Nhac lich tai kham, xu ly van de sau dich vu theo SLA.",
      "Chuyen van de vuot tham quyen cho bo phan phu trach.",
      "Tong hop insight khach hang de cai tien trai nghiem.",
      "Bao cao ket qua CSKH cho ban dieu hanh va bo phan lien quan."
    ]
  },
  "ke-toan": {
    title: "Phong Ke Toan",
    steps: [
      "Ghi nhan doanh thu, thu chi va doi soat giao dich hang ngay.",
      "Kiem tra chung tu, lap phieu thu/chi va luu tru ho so tai chinh.",
      "Xuat hoa don, theo doi cong no khach hang va doi tac.",
      "Theo doi chi phi van hanh va canh bao vuot ngan sach.",
      "Doi soat ton kho vat tu voi bo phan kho/dieu hanh.",
      "Lap bao cao tai chinh dinh ky cho ban lanh dao."
    ]
  },
  "giam-sat": {
    title: "Bo phan Giam Sat",
    steps: [
      "Theo doi muc do tuan thu quy trinh cua tung phong ban.",
      "Kiem tra chat luong van hanh va danh gia hieu qua theo KPI.",
      "Phat hien sai lech, lap bien ban va yeu cau khac phuc.",
      "Phoi hop dao tao bo sung cho cac diem yeu.",
      "Kiem tra lai sau khac phuc va xac nhan dat yeu cau.",
      "Bao cao tong hop chat luong van hanh cho giam doc."
    ]
  }
};

const POLICY_FOLDER_STRUCTURE = {
  "noi-quy-cong-ty": {
    title: "Noi quy cong ty",
    color: "#334155",
    children: {
      "hcns": ["Tuan thu gio lam viec va quy trinh xin phep theo quy dinh cong ty.", "Tat ca quyet dinh nhan su phai duoc luu vet va co phe duyet cap quan ly.", "Bao mat thong tin nhan su, ho so noi bo va du lieu luong."],
      "marketing": ["Noi dung truyen thong phai dung guideline thuong hieu da ban hanh.", "Khong cong bo du lieu noi bo khi chua co phe duyet.", "Bao cao ket qua chien dich dung ky han tuan/thang."],
      "kinh-doanh": ["Khong cam ket ngoai chinh sach da duoc cong ty phe duyet.", "Cap nhat CRM ngay sau moi lan tuong tac khach hang.", "Tuan thu quy trinh ky ket hop dong va doi soat doanh thu."],
      "dieu-hanh-ca": ["Sap xep lich theo nguyen tac uu tien khach hang va nang luc nhan su.", "Moi thay doi lich trong ngay phai cap nhat he thong ngay lap tuc.", "Khong bo qua buoc ban giao thong tin cuoi ca."],
      "dieu-duong": ["Thuc hien dung quy trinh chuyen mon va tieu chuan an toan.", "Su dung dung cu, vat tu theo dinh muc va quy dinh kho.", "Bao cao su co ngay khi phat sinh, khong xu ly vuot tham quyen."],
      "cskh": ["Phan hoi khach hang lich su toi da trong khung SLA da quy dinh.", "Khong tranh luan voi khach, uu tien xu ly theo quy trinh cong ty.", "Tat ca khieu nai phai co log va trang thai xu ly ro rang."],
      "ke-toan": ["Doi soat thu chi hang ngay, khong de chenhlech qua ngay.", "Chung tu ke toan phai day du va luu theo quy dinh.", "Xuat hoa don dung mau, dung thoi diem va dung doi tuong."],
      "giam-sat": ["Kiem tra cheo dinh ky theo checklist chung cua he thong.", "Khong bo qua vi pham quy trinh quan trong, phai lap bien ban.", "Bao cao ket qua giam sat theo dung form va dung han."]
    }
  },
  "noi-quy-phong-ban": {
    title: "Noi quy phong ban",
    color: "#0f766e",
    children: {
      "hcns": ["Hoan tat tiep nhan nhan su moi trong 48 gio lam viec.", "Danh gia thu viec theo mau chuan va luu vao ho so.", "To chuc dao tao noi bo theo ke hoach thang."],
      "marketing": ["Moi ke hoach phai co KPI ro (reach, lead, cost/lead).", "Nghiem thu noi dung truoc khi dang tai cac kenh.", "Tong hop insight thi truong va de xuat toi uu moi tuan."],
      "kinh-doanh": ["Moi lead phai duoc goi lan dau trong 2 gio lam viec.", "Cap nhat trang thai co hoi theo tung giai doan ban hang.", "Ban giao day du thong tin ca da chot cho Dieu Hanh Ca."],
      "dieu-hanh-ca": ["Lap lich theo ngay truoc 20:00 cho ngay ke tiep.", "Can doi tai nguyen de tranh qua tai nhan su.", "Tong ket ca va gui bao cao van hanh truoc 21:00."],
      "dieu-duong": ["Check thong tin khach va tien su truoc khi xu ly.", "Ghi nhat ky dich vu theo mau quy dinh ngay tai cho.", "Huong dan sau dich vu bang noi dung da duoc duyet."],
      "cskh": ["Goi cham soc sau dich vu theo moc T+1, T+3, T+7.", "Phan loai phan hoi thanh nhom: ky thuat, thai do, tai chinh.", "Chuyen cap xu ly ngay voi phan hoi muc do nghiem trong."],
      "ke-toan": ["Dong so lieu ngay truoc 10:00 sang hom sau.", "Doi soat cong no theo lich co dinh tuan/thang.", "Bao cao luu chuyen tien te theo mau chung cua cong ty."],
      "giam-sat": ["Kiem tra ngau nhien toi thieu 2 ca/ngay.", "Danh gia chat luong theo bo tieu chi KPI da phe duyet.", "Theo doi han khac phuc va xac nhan sau khac phuc."]
    }
  },
  "co-che-luong": {
    title: "Co che luong",
    color: "#7c3aed",
    children: {
      "hcns": ["Luong co ban + phu cap hanh chinh + KPI tuyen dung/dao tao.", "Danh gia KPI theo chu ky thang va quy.", "Tinh thuong theo muc do hoan thanh ke hoach nhan su."],
      "marketing": ["Luong co ban + KPI lead chat luong + KPI chi phi chuyen doi.", "Thuong them khi vuot muc lead va ty le chot.", "Dieu chinh KPI theo tung chien dich lon."],
      "kinh-doanh": ["Luong co ban + hoa hong doanh thu + thuong muc tieu.", "Co che bac thang hoa hong theo doanh so.", "Phat tru khi vi pham quy trinh ban giao/doi soat."],
      "dieu-hanh-ca": ["Luong co ban + phu cap dieu pho + KPI dung lich.", "Thuong theo ty le van hanh on dinh va giam su co.", "Danh gia theo chat luong ban giao lien bo phan."],
      "dieu-duong": ["Luong co ban + phu cap ca + KPI chat luong dich vu.", "Thuong theo muc do hai long cua khach hang.", "Cong diem theo ty le tuan thu quy trinh ky thuat."],
      "cskh": ["Luong co ban + KPI ty le phan hoi dung han.", "Thuong theo ty le khach quay lai va diem hai long.", "Danh gia theo kha nang xu ly khieu nai dung quy trinh."],
      "ke-toan": ["Luong co ban + phu cap trach nhiem tai chinh.", "Thuong theo do chinh xac doi soat va dung han bao cao.", "Co che phat khi sai lech so lieu nghiem trong."],
      "giam-sat": ["Luong co ban + KPI chat luong he thong.", "Thuong theo ty le khac phuc sau giam sat.", "Danh gia theo hieu qua cai tien quy trinh toan bo." ]
    }
  },
  "thuong-phat": {
    title: "Quy dinh thuong phat",
    color: "#be123c",
    children: {
      "hcns": ["Thuong khi hoan thanh tuyen dung dung han va dung chat luong.", "Phat khi chua cap nhat day du ho so lao dong bat buoc.", "Nhac nho/canh cao khi tre han bao cao dinh ky."],
      "marketing": ["Thuong khi vuot KPI lead chat luong theo thang.", "Phat khi noi dung sai guideline hoac dang sai kenh.", "Canh cao khi bao cao khong trung thuc so lieu."],
      "kinh-doanh": ["Thuong theo moc doanh so va ty le chot.", "Phat khi thong tin cam ket voi khach sai quy dinh.", "Canh cao khi bo sot cap nhat CRM lien tiep."],
      "dieu-hanh-ca": ["Thuong khi van hanh on dinh, khong co su co nghiem trong.", "Phat khi dieu pho gay qua tai hoac tre lich lặp lai.", "Canh cao khi ban giao thong tin thieu gay anh huong dich vu."],
      "dieu-duong": ["Thuong theo diem danh gia chat luong sau dich vu.", "Phat khi vi pham quy trinh an toan/chuyen mon.", "Tam dung ca truc neu vi pham nghiem trong lap lai."],
      "cskh": ["Thuong khi duy tri ty le hai long va quay lai cao.", "Phat khi tre SLA xu ly khieu nai khong ly do.", "Canh cao khi giao tiep khong dung quy tac ung xu."],
      "ke-toan": ["Thuong khi doi soat chinh xac, khong phat sinh sai lech.", "Phat khi sai pham chung tu/hoa don theo muc do.", "Canh cao khi tre han nop bao cao tai chinh."],
      "giam-sat": ["Thuong khi de xuat cai tien duoc ap dung hieu qua.", "Phat khi bo sot vi pham he thong nghiem trong.", "Canh cao khi bao cao giam sat khong day du bang chung."]
    }
  }
};

const POLICY_DEPARTMENT_LABELS = {
  "hcns": "Phong HCNS",
  "marketing": "Phong Marketing",
  "kinh-doanh": "Phong Kinh doanh",
  "dieu-hanh-ca": "Bo phan Dieu hanh ca",
  "dieu-duong": "Bo phan Dieu duong",
  "cskh": "Bo phan Cham soc khach hang",
  "ke-toan": "Phong Ke toan",
  "giam-sat": "Bo phan Giam sat"
};

function loadJSON(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) {
      localStorage.setItem(key, JSON.stringify(fallback));
      return fallback;
    }
    return JSON.parse(raw);
  } catch {
    return fallback;
  }
}

function saveJSON(key, value) {
  localStorage.setItem(key, JSON.stringify(value));
  queueCriticalStateSync(key);
}

function normalizeDataSourceConfig(rawConfig = {}) {
  const type = rawConfig.type === "api" || rawConfig.type === "sheet" ? rawConfig.type : "local";
  return {
    type,
    url: String(rawConfig.url || "").trim()
  };
}

function clonePlain(value) {
  return value == null ? value : JSON.parse(JSON.stringify(value));
}

function normalizeLoginPrefs(rawPrefs = {}) {
  const remember = Boolean(rawPrefs.remember);
  return {
    remember,
    username: remember ? String(rawPrefs.username || "") : "",
    password: remember ? String(rawPrefs.password || "") : ""
  };
}

function getLoginPrefs() {
  return normalizeLoginPrefs(loadJSON(STORAGE.loginPrefs, { remember: false, username: "", password: "" }));
}

function saveLoginPrefs(remember, username = "", password = "") {
  const payload = normalizeLoginPrefs({ remember, username, password });
  saveJSON(STORAGE.loginPrefs, payload);
}

function applyLoginPrefsToForm() {
  const prefs = getLoginPrefs();
  if (els.loginRemember) els.loginRemember.checked = prefs.remember;
  if (els.loginUsername) els.loginUsername.value = prefs.username;
  if (els.loginPassword) {
    els.loginPassword.value = prefs.password;
    els.loginPassword.type = "password";
  }
  if (els.toggleLoginPasswordBtn) {
    els.toggleLoginPasswordBtn.textContent = "👁";
    els.toggleLoginPasswordBtn.setAttribute("aria-label", "Hiện mật khẩu");
  }
}

function normalizeCustomer(customer) {
  return {
    ...customer,
    status: customer.status || CUSTOMER_STATUSES[0],
    owner: customer.owner || customer.assignee || "Chưa gán",
    source: customer.source || "Nhập tay",
    updatedAt: Number(customer.updatedAt) || Date.now()
  };
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function renderAddrDistricts(province) {
  els.customerAddrDistrict.innerHTML = '<option value="">Lựa chọn</option>';
  els.customerAddrWard.innerHTML = '<option value="">Lựa chọn</option>';
  els.customerAddrWard.disabled = true;
  if (!province || !VN_ADDRESSES[province]) {
    els.customerAddrDistrict.disabled = true;
    return;
  }
  const districts = Object.keys(VN_ADDRESSES[province]);
  els.customerAddrDistrict.innerHTML =
    '<option value="">Lựa chọn</option>' +
    districts.map((d) => `<option value="${d}">${d}</option>`).join("");
  els.customerAddrDistrict.disabled = false;
}

function renderAddrWards(province, district) {
  els.customerAddrWard.innerHTML = '<option value="">Lựa chọn</option>';
  els.customerAddrWard.disabled = true;
  if (!province || !district || !VN_ADDRESSES[province]?.[district]?.length) return;
  const wards = VN_ADDRESSES[province][district];
  els.customerAddrWard.innerHTML =
    '<option value="">Lựa chọn</option>' +
    wards.map((w) => `<option value="${w}">${w}</option>`).join("");
  els.customerAddrWard.disabled = false;
}

function composeAddressFromSelects() {
  const ward = els.customerAddrWard.value;
  const district = els.customerAddrDistrict.value;
  const province = els.customerAddrProvince.value;
  const parts = [ward, district, province].filter(Boolean);
  if (parts.length > 0) els.customerAddress.value = parts.join(", ");
}

function normalizeCustomerFilterState(rawState = {}) {
  const legacyStart = rawState.start || "";
  const legacyEnd = rawState.end || "";
  const date = rawState.date || "";
  const resolvedStart = legacyStart || date;
  const resolvedEnd = legacyEnd || date;

  let start = resolvedStart;
  let end = resolvedEnd;
  if (start && end && start > end) {
    const temp = start;
    start = end;
    end = temp;
  }

  return {
    start,
    end,
    owner: rawState.owner || "all",
    status: rawState.status || "all",
    source: rawState.source || "all",
    keyword: rawState.keyword || ""
  };
}

function saveCustomerFilterState() {
  saveJSON(STORAGE.customerFilters, customerFilterState);
}

function normalizeRolePermissions(rawPermissions = {}) {
  const permissions = {};
  Object.entries(ROLES).forEach(([roleKey, base]) => {
    const incoming = rawPermissions[roleKey] || {};
    const merged = {
      ...base,
      ...incoming,
      pageAccess: Array.isArray(incoming.pageAccess) ? incoming.pageAccess : base.pageAccess
    };

    merged.canViewData = Boolean(merged.canViewData);
    merged.canViewUsers = Boolean(merged.canViewUsers);
    merged.canSubmitReport = Boolean(merged.canSubmitReport);
    merged.canManageUsers = Boolean(merged.canManageUsers);
    merged.canSyncData = Boolean(merged.canSyncData);
    merged.canExportPdf = Boolean(merged.canExportPdf);
    merged.pageAccess = Array.from(new Set((merged.pageAccess || []).filter((key) => APP_PAGE_KEYS.includes(key))));

    // Roles that can sync data should always have access to the data source page.
    if (merged.canSyncData && !merged.pageAccess.includes("access")) {
      merged.pageAccess.push("access");
    }

    if (roleKey === "admin" || roleKey === "ceo") {
      merged.canViewData = true;
      merged.canViewUsers = true;
      merged.canSubmitReport = true;
      merged.canManageUsers = true;
      merged.canSyncData = true;
      merged.canExportPdf = true;
      merged.pageAccess = [...APP_PAGE_KEYS];
    }

    permissions[roleKey] = merged;
  });
  return permissions;
}

function saveRolePermissionsState() {
  saveJSON(STORAGE.rolePermissions, rolePermissionsState);
}

function canAccessPage(pageKey) {
  if (!authState.loggedIn) return false;
  if (pageKey === "home") return true;
  const user = getCurrentUser();
  if (!user) return false;
  const permissions = getRolePermissions(user.roleKey);
  return (permissions.pageAccess || []).includes(pageKey);
}

function renderPermissionModal() {
  if (!els.permissionRoleSelect || !els.permissionFeatureGrid || !els.permissionPageGrid) return;
  const roleKey = permissionEditingRole;
  const rolePerms = rolePermissionsState[roleKey] || getRolePermissions(roleKey);
  const lockFull = roleKey === "admin" || roleKey === "ceo";
  const disabledAttr = lockFull ? "disabled" : "";

  const featureMeta = [
    ["canViewData", "Xem dữ liệu"],
    ["canViewUsers", "Xem nhân sự"],
    ["canSubmitReport", "Nhập/Cập nhật dữ liệu"],
    ["canSyncData", "Đồng bộ dữ liệu"],
    ["canExportPdf", "Xuất PDF"],
    ["canManageUsers", "Quản lý nhân sự"]
  ];

  els.permissionRoleSelect.value = roleKey;
  els.permissionFeatureGrid.innerHTML = featureMeta
    .map(([key, label]) => `
      <label class="metrics-dept-card" style="display:flex;align-items:center;gap:8px;cursor:pointer;">
        <input type="checkbox" data-perm-feature="${key}" ${rolePerms[key] ? "checked" : ""} ${disabledAttr} />
        <span>${label}</span>
      </label>
    `)
    .join("");

  els.permissionPageGrid.innerHTML = APP_PAGE_KEYS
    .map((pageKey) => `
      <label class="metrics-dept-card" style="display:flex;align-items:center;gap:8px;cursor:pointer;">
        <input type="checkbox" data-perm-page="${pageKey}" ${(rolePerms.pageAccess || []).includes(pageKey) ? "checked" : ""} ${disabledAttr} />
        <span>${APP_PAGE_LABELS[pageKey] || pageKey}</span>
      </label>
    `)
    .join("");

  if (els.permissionModalStatus) {
    els.permissionModalStatus.textContent = lockFull
      ? "Admin/CEO luôn full tính năng và toàn bộ trang."
      : "Tùy chỉnh chức năng và trang xem cho vai trò đang chọn.";
  }
}

function openPermissionModal() {
  permissionEditingRole = "staff";
  renderPermissionModal();
  if (els.permissionModal) els.permissionModal.classList.remove("hidden");
}

function closePermissionModal() {
  if (els.permissionModal) els.permissionModal.classList.add("hidden");
  if (els.permissionModalStatus) els.permissionModalStatus.textContent = "";
}

function normalizeCustomerCareFilterState(rawState = {}) {
  let start = rawState.start || "";
  let end = rawState.end || "";
  if (start && end && start > end) {
    const temp = start;
    start = end;
    end = temp;
  }
  return {
    start,
    end,
    staff: rawState.staff || "all",
    status: rawState.status || "all",
    source: rawState.source || "all",
    progress: rawState.progress || "all",
    keyword: rawState.keyword || ""
  };
}

function saveCustomerCareFilterState() {
  saveJSON(STORAGE.customerCareFilters, customerCareFilterState);
}

function showToast(message, type = "success") {
  const container = document.getElementById("toastContainer");
  if (!container) return;
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  const icon = { success: "✅", error: "❌", warning: "⚠️", info: "ℹ️" }[type] || "✅";
  toast.innerHTML = `<span>${icon}</span><span>${message}</span>`;
  container.appendChild(toast);
  setTimeout(() => toast.remove(), 3000);
}

function getInitials(name = "") {
  const parts = String(name).trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "NC";
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return `${parts[0][0] || ""}${parts[parts.length - 1][0] || ""}`.toUpperCase();
}

function formatNewsDateTime(timestamp) {
  return new Date(timestamp).toLocaleString("vi-VN", {
    hour12: false,
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function formatNewsEventDate(dateValue) {
  if (!dateValue) return "--/--";
  const [y, m, d] = String(dateValue).split("-");
  if (!y || !m || !d) return "--/--";
  return `${d}/${m}`;
}

function getNewsEventVoteSummary(eventItem) {
  const votes = eventItem?.votes && typeof eventItem.votes === "object" ? Object.values(eventItem.votes) : [];
  const yes = votes.filter((value) => value === "yes").length;
  const no = votes.filter((value) => value === "no").length;
  return { yes, no, total: votes.length };
}

function formatFileSize(bytes) {
  const size = Math.max(0, Number(bytes) || 0);
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function cloneAttachmentList(list = []) {
  return Array.isArray(list)
    ? list.map((item) => ({
      id: item.id || `nfa-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      name: item.name || "Tệp đính kèm",
      type: item.type || "application/octet-stream",
      data: item.data || "",
      size: Math.max(0, Number(item.size) || 0)
    })).filter((item) => item.data)
    : [];
}

function readFilesAsAttachments(files = []) {
  return Promise.all(
    Array.from(files).map((file) => new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        const data = typeof event.target?.result === "string" ? event.target.result : "";
        resolve({
          id: `nfa-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
          name: file.name,
          type: file.type || "application/octet-stream",
          data,
          size: file.size || 0
        });
      };
      reader.readAsDataURL(file);
    }))
  );
}

function renderEditorAttachmentList(container, list = [], context = "post") {
  if (!container) return;
  if (!list.length) {
    container.innerHTML = '<p class="muted">Chưa có tệp đính kèm.</p>';
    return;
  }
  container.innerHTML = list
    .map((item, index) => `
      <div class="news-editor-attachment-item">
        <a href="${item.data}" download="${escapeHtml(item.name)}">${escapeHtml(item.name)}</a>
        <span class="muted">${escapeHtml(formatFileSize(item.size))}</span>
        <button type="button" class="btn secondary" data-news-remove-attachment="${context}" data-news-attachment-index="${index}">Xóa</button>
      </div>
    `)
    .join("");
}

function isImageAttachment(item) {
  return String(item?.type || "").startsWith("image/") || String(item?.data || "").startsWith("data:image/");
}

function renderNewsAttachmentFiles(list = [], contentType = "post", contentId = "") {
  if (!list.length) return "";
  const images = list.filter((item) => isImageAttachment(item));
  const files = list.filter((item) => !isImageAttachment(item));
  return `
    <div class="news-attachment-files">
      ${images.map((item, index) => `
        <button type="button" class="news-attachment-image-btn" data-news-open-detail-type="${contentType}" data-news-open-detail-id="${contentId}" data-news-detail-image-index="${index}">
          <img src="${item.data}" alt="${escapeHtml(item.name)}" class="news-attachment-image" />
        </button>
      `).join("")}
      ${files.map((item) => `
        <a class="news-attachment-file" href="${item.data}" target="_blank" rel="noreferrer">
          📎 ${escapeHtml(item.name)}
        </a>
      `).join("")}
    </div>
  `;
}

function getNewsContentByType(type, id) {
  if (type === "event") return newsEvents.find((item) => item.id === id) || null;
  return newsPosts.find((item) => item.id === id) || null;
}

function openNewsContentDetail(type, id, imageIndex = 0) {
  if (!els.newsContentDetailModal || !els.newsContentDetailTitle || !els.newsContentDetailMeta || !els.newsContentDetailText || !els.newsContentDetailAttachments || !els.newsContentDetailImagePreview) return;
  const item = getNewsContentByType(type, id);
  if (!item) return;

  activeNewsDetail = { type, id, imageIndex: Math.max(0, Number(imageIndex) || 0) };
  const title = type === "event" ? item.title : item.authorName;
  const meta = type === "event"
    ? `${formatNewsEventDate(item.date)} ${item.time || "--:--"} | ${item.location || "Chưa có địa điểm"}`
    : `${formatNewsDateTime(item.createdAt)} | ${item.department || "Toàn hệ thống"}`;
  const text = type === "event" ? (item.description || "Không có ghi chú") : (item.content || "");

  els.newsContentDetailTitle.textContent = title;
  els.newsContentDetailMeta.textContent = meta;
  els.newsContentDetailText.textContent = text;

  const attachments = item.attachments || [];
  const images = attachments.filter((entry) => isImageAttachment(entry));
  const files = attachments.filter((entry) => !isImageAttachment(entry));

  if (images.length) {
    const img = images[Math.min(activeNewsDetail.imageIndex, images.length - 1)];
    els.newsContentDetailImagePreview.innerHTML = `<img src="${img.data}" alt="${escapeHtml(img.name)}" />`;
    els.newsContentDetailImagePreview.classList.remove("hidden");
  } else {
    els.newsContentDetailImagePreview.innerHTML = "";
    els.newsContentDetailImagePreview.classList.add("hidden");
  }

  els.newsContentDetailAttachments.innerHTML = attachments.length
    ? `
      <div class="news-detail-file-list">
        ${images.map((entry, index) => `<button type="button" class="news-attachment-file" data-news-open-detail-type="${type}" data-news-open-detail-id="${id}" data-news-detail-image-index="${index}">🖼 ${escapeHtml(entry.name)}</button>`).join("")}
        ${files.map((entry) => `<a class="news-attachment-file" href="${entry.data}" target="_blank" rel="noreferrer">📎 ${escapeHtml(entry.name)}</a>`).join("")}
      </div>
      ${files.length ? `<p class="muted" style="margin-top:6px;">Tệp tài liệu: ${files.length}</p>` : ""}
    `
    : '<p class="muted">Không có tệp đính kèm.</p>';

  els.newsContentDetailModal.classList.remove("hidden");
}

function closeNewsContentDetail() {
  if (!els.newsContentDetailModal) return;
  els.newsContentDetailModal.classList.add("hidden");
  activeNewsDetail = null;
}

function getVoteUsersByChoice(eventItem, choice) {
  const targetChoice = choice === "no" ? "no" : "yes";
  const votes = eventItem?.votes && typeof eventItem.votes === "object" ? eventItem.votes : {};
  return Object.entries(votes)
    .filter(([, value]) => value === targetChoice)
    .map(([userId]) => {
      const user = users.find((item) => item.id === userId);
      if (!user) return { id: userId, name: `User ${userId}` };
      return { id: userId, name: user.fullName || user.username || userId };
    });
}

function renderNewsPosts() {
  if (!els.newsPostList) return;
  if (!newsPosts.length) {
    els.newsPostList.innerHTML = '<article class="news-post card"><p class="muted">Chưa có bài đăng. Hãy tạo thông báo đầu tiên.</p></article>';
    return;
  }

  els.newsPostList.innerHTML = newsPosts
    .slice(0, 30)
    .map((post) => {
      const toneClass = post.tone === "warn" ? "news-tag-warn" : post.tone === "good" ? "news-tag-good" : "";
      const tags = post.tags.length ? post.tags : [post.department || "Nội bộ"];
      return `
        <article class="news-post card">
          <header>
            <div class="news-post-head-main">
              <div class="news-avatar">${escapeHtml(getInitials(post.authorName))}</div>
              <div>
                <h4>${escapeHtml(post.authorName)}</h4>
                <p class="muted">${escapeHtml(formatNewsDateTime(post.createdAt))}</p>
              </div>
            </div>
            <div class="news-card-actions">
              <button type="button" class="btn secondary user-action-toggle news-action-toggle" data-news-post-toggle="${post.id}" title="Thao tác">...</button>
              <div class="user-action-menu hidden" data-news-post-menu="${post.id}">
                <button class="user-action-item" type="button" data-news-post-edit="${post.id}">✏️ Sửa</button>
                <button class="user-action-item user-action-item--danger" type="button" data-news-post-delete="${post.id}">🗑 Xóa</button>
              </div>
            </div>
          </header>
          <p>${escapeHtml(post.content)}</p>
          ${renderNewsAttachmentFiles(post.attachments || [], "post", post.id)}
          <div class="news-tag-row">
            ${tags.map((tag, idx) => `<span class="news-tag ${idx === 0 ? toneClass : ""}">${escapeHtml(tag)}</span>`).join("")}
          </div>
          <footer>
            <span>${post.views.toLocaleString("vi-VN")} lượt xem</span>
            <span>${post.comments.toLocaleString("vi-VN")} bình luận</span>
          </footer>
        </article>
      `;
    })
    .join("");
}

function renderNewsPinned() {
  if (!els.newsPinnedList) return;
  const sorted = [...newsPinned].sort((a, b) => b.createdAt - a.createdAt).slice(0, 8);
  if (!sorted.length) {
    els.newsPinnedList.innerHTML = "<li>Chưa có thông báo ghim.</li>";
    return;
  }
  els.newsPinnedList.innerHTML = sorted.map((item) => `<li>${escapeHtml(item.text)}</li>`).join("");
}

function renderNewsEvents() {
  if (!els.newsEventsList) return;
  const sorted = [...newsEvents].sort((a, b) => a.date.localeCompare(b.date)).slice(0, 8);
  if (!sorted.length) {
    els.newsEventsList.innerHTML = '<p class="muted">Chưa có sự kiện.</p>';
    return;
  }
  els.newsEventsList.innerHTML = sorted
    .map((item) => `
      <div class="news-event-item">
        <strong>${escapeHtml(formatNewsEventDate(item.date))}</strong>
        <span>
          ${escapeHtml(item.title)}
          <small class="muted news-event-meta">${escapeHtml(item.time || "--:--")} • ${escapeHtml(item.location || "Chưa có địa điểm")}</small>
        </span>
      </div>
    `)
    .join("");
}

function renderNewsEventFeed() {
  if (!els.newsEventFeedList) return;
  const sorted = [...newsEvents].sort((a, b) => `${a.date} ${a.time || ""}`.localeCompare(`${b.date} ${b.time || ""}`));
  if (!sorted.length) {
    els.newsEventFeedList.innerHTML = "";
    return;
  }

  const currentUserId = getCurrentUser()?.id || "";
  els.newsEventFeedList.innerHTML = sorted
    .slice(0, 6)
    .map((eventItem) => {
      const vote = currentUserId ? eventItem.votes?.[currentUserId] : "";
      const summary = getNewsEventVoteSummary(eventItem);
      const voteStatusText = vote === "yes" ? "Bạn đã chọn: Tham gia" : vote === "no" ? "Bạn đã chọn: Vắng mặt" : "Bạn chưa xác nhận";
      return `
        <article class="card news-event-feed-card">
          <header>
            <div>
              <strong>${escapeHtml(eventItem.title)}</strong>
              <span class="muted" style="display:block;">${escapeHtml(formatNewsEventDate(eventItem.date))} • ${escapeHtml(eventItem.time || "--:--")}</span>
            </div>
            <div class="news-card-actions">
              <button type="button" class="btn secondary user-action-toggle news-action-toggle" data-news-event-toggle="${eventItem.id}" title="Thao tác">...</button>
              <div class="user-action-menu hidden" data-news-event-menu="${eventItem.id}">
                <button class="user-action-item" type="button" data-news-event-edit="${eventItem.id}">✏️ Sửa</button>
                <button class="user-action-item user-action-item--danger" type="button" data-news-event-delete="${eventItem.id}">🗑 Xóa</button>
              </div>
            </div>
          </header>
          <p class="muted">${escapeHtml(eventItem.location || "Chưa có địa điểm")}${eventItem.description ? ` | ${escapeHtml(eventItem.description)}` : ""}</p>
          ${renderNewsAttachmentFiles(eventItem.attachments || [], "event", eventItem.id)}
          <div class="news-event-vote-row">
            <button type="button" class="btn secondary ${vote === "yes" ? "news-vote-active" : ""}" data-news-vote-event="${eventItem.id}" data-news-vote-choice="yes">Tham gia</button>
            <button type="button" class="btn secondary ${vote === "no" ? "news-vote-active" : ""}" data-news-vote-event="${eventItem.id}" data-news-vote-choice="no">Vắng mặt</button>
            <span class="muted">${escapeHtml(voteStatusText)}</span>
            <button type="button" class="btn secondary" data-news-vote-detail-event="${eventItem.id}" data-news-vote-detail-choice="yes">Có: ${summary.yes}</button>
            <button type="button" class="btn secondary" data-news-vote-detail-event="${eventItem.id}" data-news-vote-detail-choice="no">Vắng: ${summary.no}</button>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderNewsTrends() {
  if (!els.newsTrendList) return;
  const topicCount = new Map();
  newsPosts.forEach((post) => {
    (post.tags || []).forEach((tag) => {
      const normalized = String(tag || "").trim().toLowerCase();
      if (!normalized) return;
      topicCount.set(normalized, (topicCount.get(normalized) || 0) + 1);
    });
  });
  const trends = [...topicCount.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 6)
    .map(([tag]) => `#${tag.replace(/\s+/g, "-")}`);

  if (!trends.length) {
    els.newsTrendList.innerHTML = '<p class="muted">#bang-tin-noi-bo</p>';
    return;
  }

  els.newsTrendList.innerHTML = trends.map((tag) => `<p class="muted">${escapeHtml(tag)}</p>`).join("");
}

function renderNewsPage() {
  renderNewsPosts();
  renderNewsEventFeed();
  renderNewsPinned();
  renderNewsEvents();
  renderNewsTrends();
}

function openNewsComposer() {
  if (!els.newsComposerEditor) return;
  els.newsComposerEditor.classList.remove("hidden");
  renderEditorAttachmentList(els.newsComposerAttachmentList, pendingNewsPostAttachments, "post");
  if (els.newsComposerText) els.newsComposerText.focus();
}

function resetNewsComposerEditingState() {
  editingNewsPostId = null;
  pendingNewsPostAttachments = [];
  if (els.newsComposerAttachmentInput) els.newsComposerAttachmentInput.value = "";
  renderEditorAttachmentList(els.newsComposerAttachmentList, pendingNewsPostAttachments, "post");
  if (els.newsSubmitPostBtn) els.newsSubmitPostBtn.textContent = "Đăng thông báo";
}

function openNewsEventModal() {
  if (!els.newsEventModal) return;
  editingNewsEventId = null;
  pendingNewsEventAttachments = [];
  if (els.newsEventAttachmentInput) els.newsEventAttachmentInput.value = "";
  renderEditorAttachmentList(els.newsEventAttachmentList, pendingNewsEventAttachments, "event");
  if (els.saveNewsEventBtn) els.saveNewsEventBtn.textContent = "Lưu sự kiện";
  if (els.newsEventTitle && els.newsComposerText) {
    const seedTitle = els.newsComposerText.value.trim();
    els.newsEventTitle.value = seedTitle;
  }
  if (els.newsEventDate && els.newsComposerEventDate) {
    els.newsEventDate.value = els.newsComposerEventDate.value || today;
  }
  if (els.newsEventTime && !els.newsEventTime.value) {
    els.newsEventTime.value = "09:00";
  }
  if (els.newsEventLocation && !els.newsEventLocation.value) {
    els.newsEventLocation.value = "Phòng họp chính";
  }
  if (els.newsEventModalStatus) els.newsEventModalStatus.textContent = "";
  els.newsEventModal.classList.remove("hidden");
}

function closeNewsEventModal() {
  if (!els.newsEventModal) return;
  els.newsEventModal.classList.add("hidden");
  editingNewsEventId = null;
  pendingNewsEventAttachments = [];
  if (els.newsEventAttachmentInput) els.newsEventAttachmentInput.value = "";
  renderEditorAttachmentList(els.newsEventAttachmentList, pendingNewsEventAttachments, "event");
  if (els.saveNewsEventBtn) els.saveNewsEventBtn.textContent = "Lưu sự kiện";
  if (els.newsEventModalStatus) els.newsEventModalStatus.textContent = "";
}

function openNewsVoteDetail(eventId, choice) {
  if (!els.newsVoteDetailModal || !els.newsVoteDetailTitle || !els.newsVoteDetailList) return;
  const eventItem = newsEvents.find((item) => item.id === eventId);
  if (!eventItem) return;
  const normalizedChoice = choice === "no" ? "no" : "yes";
  const list = getVoteUsersByChoice(eventItem, normalizedChoice);
  const label = normalizedChoice === "yes" ? "Tham gia" : "Vắng mặt";
  els.newsVoteDetailTitle.textContent = `${eventItem.title} - Danh sách ${label}`;
  els.newsVoteDetailList.innerHTML = list.length
    ? `<ul>${list.map((item) => `<li>${escapeHtml(item.name)}</li>`).join("")}</ul>`
    : `<p class="muted">Chưa có người chọn ${label.toLowerCase()}.</p>`;
  els.newsVoteDetailModal.classList.remove("hidden");
}

function closeNewsVoteDetail() {
  if (!els.newsVoteDetailModal) return;
  els.newsVoteDetailModal.classList.add("hidden");
}

function submitNewsPost() {
  if (!authState.loggedIn) {
    showToast("Bạn cần đăng nhập để đăng thông báo.", "warning");
    return;
  }
  if (!els.newsComposerText || !els.newsComposerDept || !els.newsComposerImportant) return;

  const content = els.newsComposerText.value.trim();
  const department = els.newsComposerDept.value || "Toàn hệ thống";
  const important = els.newsComposerImportant.checked;
  if (!content) {
    showToast("Vui lòng nhập nội dung thông báo.", "warning");
    return;
  }

  const user = getCurrentUser();
  const authorName = user?.fullName || user?.username || "Hệ thống";
  const tags = important ? ["Quan trọng", department] : [department, "Nội bộ"];
  const attachments = cloneAttachmentList(pendingNewsPostAttachments);
  const previousPost = editingNewsPostId ? clonePlain(newsPosts.find((item) => item.id === editingNewsPostId) || null) : null;
  if (editingNewsPostId) {
    newsPosts = newsPosts.map((item) => {
      if (item.id !== editingNewsPostId) return item;
      return {
        ...item,
        authorName,
        department,
        content,
        tags,
        attachments,
        tone: important ? "warn" : "default"
      };
    });
  } else {
    newsPosts.unshift({
      id: `np-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      authorName,
      department,
      createdAt: Date.now(),
      content,
      tags,
      attachments,
      views: 0,
      comments: 0,
      tone: important ? "warn" : "default"
    });
  }

  if (important) {
    newsPinned.unshift({
      id: `pin-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      text: `${department}: ${content.slice(0, 100)}${content.length > 100 ? "..." : ""}`,
      createdAt: Date.now()
    });
    if (newsPinned.length > 20) newsPinned = newsPinned.slice(0, 20);
    saveJSON(STORAGE.newsPinned, newsPinned);
  }

  saveJSON(STORAGE.newsPosts, newsPosts);
  renderNewsPage();
  els.newsComposerText.value = "";
  els.newsComposerImportant.checked = false;
  showToast(editingNewsPostId ? "Đã cập nhật thông báo." : "Đã đăng thông báo lên Bảng tin.");
  logActivity(
    "Bảng tin",
    editingNewsPostId ? "Cập nhật thông báo" : "Đăng thông báo",
    `${department} | ${content.slice(0, 80)}`,
    editingNewsPostId && previousPost ? { restoreAction: { kind: "news-post-edit", previousPost } } : {}
  );
  resetNewsComposerEditingState();
}

function createNewsEvent() {
  openNewsEventModal();
}

function saveNewsEvent() {
  if (!authState.loggedIn) {
    showToast("Bạn cần đăng nhập để tạo sự kiện.", "warning");
    return;
  }
  if (!els.newsEventTitle || !els.newsEventDate || !els.newsEventTime || !els.newsEventLocation || !els.newsEventDescription) return;

  const title = els.newsEventTitle.value.trim();
  const date = els.newsEventDate.value || today;
  const time = els.newsEventTime.value || "09:00";
  const location = els.newsEventLocation.value.trim();
  const description = els.newsEventDescription.value.trim();

  if (!title || !location) {
    if (els.newsEventModalStatus) {
      els.newsEventModalStatus.textContent = "Vui lòng nhập tiêu đề và địa điểm sự kiện.";
    }
    return;
  }

  const user = getCurrentUser();
  const voterId = user?.id || "";
  const attachments = cloneAttachmentList(pendingNewsEventAttachments);
  const previousEvent = editingNewsEventId ? clonePlain(newsEvents.find((item) => item.id === editingNewsEventId) || null) : null;
  if (editingNewsEventId) {
    newsEvents = newsEvents.map((item) => {
      if (item.id !== editingNewsEventId) return item;
      return {
        ...item,
        date,
        time,
        title,
        location,
        description,
        attachments
      };
    });
  } else {
    const votes = voterId ? { [voterId]: "yes" } : {};
    newsEvents.unshift({
      id: `ne-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      date,
      time,
      title,
      location,
      description,
      attachments,
      votes,
      createdAt: Date.now()
    });
  }

  saveJSON(STORAGE.newsEvents, newsEvents);
  renderNewsPage();
  if (els.newsComposerEventDate) els.newsComposerEventDate.value = date;
  if (els.newsEventTitle) els.newsEventTitle.value = "";
  if (els.newsEventLocation) els.newsEventLocation.value = "";
  if (els.newsEventDescription) els.newsEventDescription.value = "";
  if (els.newsEventModalStatus) els.newsEventModalStatus.textContent = "";
  showToast(editingNewsEventId ? "Đã cập nhật sự kiện." : "Đã tạo sự kiện và lịch hẹn.");
  logActivity(
    "Bảng tin",
    editingNewsEventId ? "Cập nhật sự kiện" : "Tạo sự kiện",
    `${date} ${time} | ${title.slice(0, 80)}`,
    editingNewsEventId && previousEvent ? { restoreAction: { kind: "news-event-edit", previousEvent } } : {}
  );
  closeNewsEventModal();
  editingNewsEventId = null;
  if (els.saveNewsEventBtn) els.saveNewsEventBtn.textContent = "Lưu sự kiện";
}

function voteNewsEvent(eventId, choice) {
  const user = getCurrentUser();
  if (!user) {
    showToast("Bạn cần đăng nhập để vote tham gia.", "warning");
    return;
  }
  const validChoice = choice === "no" ? "no" : "yes";
  newsEvents = newsEvents.map((eventItem) => {
    if (eventItem.id !== eventId) return eventItem;
    return {
      ...eventItem,
      votes: {
        ...(eventItem.votes || {}),
        [user.id]: validChoice
      }
    };
  });
  saveJSON(STORAGE.newsEvents, newsEvents);
  renderNewsPage();
  showToast(validChoice === "yes" ? "Đã xác nhận tham gia." : "Đã cập nhật không tham gia.", "info");
}

function editNewsPost(postId) {
  const post = newsPosts.find((item) => item.id === postId);
  if (!post || !els.newsComposerText || !els.newsComposerDept) return;
  openNewsComposer();
  editingNewsPostId = postId;
  els.newsComposerText.value = post.content || "";
  els.newsComposerDept.value = post.department || "Toàn hệ thống";
  pendingNewsPostAttachments = cloneAttachmentList(post.attachments || []);
  renderEditorAttachmentList(els.newsComposerAttachmentList, pendingNewsPostAttachments, "post");
  if (els.newsComposerImportant) els.newsComposerImportant.checked = post.tone === "warn";
  if (els.newsSubmitPostBtn) els.newsSubmitPostBtn.textContent = "Cập nhật thông báo";
  hideAllActionMenus();
}

function deleteNewsPost(postId) {
  const target = newsPosts.find((item) => item.id === postId);
  if (!target) return;
  newsPosts = newsPosts.filter((item) => item.id !== postId);
  saveJSON(STORAGE.newsPosts, newsPosts);
  renderNewsPage();
  hideAllActionMenus();
  showToast("Đã xóa bài đăng.", "info");
  logActivity("Bảng tin", "Xóa thông báo", target.content.slice(0, 80), { restoreAction: { kind: "news-post-delete", deletedPost: clonePlain(target) } });
}

function editNewsEvent(eventId) {
  const eventItem = newsEvents.find((item) => item.id === eventId);
  if (!eventItem || !els.newsEventTitle || !els.newsEventDate || !els.newsEventTime || !els.newsEventLocation || !els.newsEventDescription) return;
  editingNewsEventId = eventId;
  els.newsEventTitle.value = eventItem.title || "";
  els.newsEventDate.value = eventItem.date || today;
  els.newsEventTime.value = eventItem.time || "09:00";
  els.newsEventLocation.value = eventItem.location || "";
  els.newsEventDescription.value = eventItem.description || "";
  pendingNewsEventAttachments = cloneAttachmentList(eventItem.attachments || []);
  renderEditorAttachmentList(els.newsEventAttachmentList, pendingNewsEventAttachments, "event");
  if (els.newsEventAttachmentInput) els.newsEventAttachmentInput.value = "";
  if (els.saveNewsEventBtn) els.saveNewsEventBtn.textContent = "Cập nhật sự kiện";
  if (els.newsEventModalStatus) els.newsEventModalStatus.textContent = "";
  els.newsEventModal?.classList.remove("hidden");
  hideAllActionMenus();
}

function deleteNewsEvent(eventId) {
  const target = newsEvents.find((item) => item.id === eventId);
  if (!target) return;
  newsEvents = newsEvents.filter((item) => item.id !== eventId);
  saveJSON(STORAGE.newsEvents, newsEvents);
  renderNewsPage();
  hideAllActionMenus();
  showToast("Đã xóa sự kiện.", "info");
  logActivity("Bảng tin", "Xóa sự kiện", target.title, { restoreAction: { kind: "news-event-delete", deletedEvent: clonePlain(target) } });
}

function attachNewsDepartment() {
  if (!els.newsComposerText || !els.newsComposerDept) return;
  const department = els.newsComposerDept.value || "Toàn hệ thống";
  const marker = `#${department.toLowerCase().replace(/\s+/g, "-")}`;
  const text = els.newsComposerText.value.trim();
  els.newsComposerText.value = text.includes(marker) ? text : `${text}${text ? " " : ""}${marker}`;
  showToast(`Đã gắn phòng ban: ${department}.`, "info");
}

function hideAllActionMenus() {
  document.querySelectorAll(".user-action-menu").forEach((menu) => {
    menu.classList.add("hidden");
    menu.classList.remove("user-action-menu-floating");
    menu.style.left = "";
    menu.style.top = "";
  });
  document.querySelectorAll("td.action-menu-open-cell").forEach((cell) => cell.classList.remove("action-menu-open-cell"));
}

function openActionMenuAtToggle(toggleBtn, menuEl) {
  if (!(toggleBtn instanceof HTMLElement) || !(menuEl instanceof HTMLElement)) return;
  hideAllActionMenus();
  menuEl.classList.remove("hidden");
  menuEl.classList.add("user-action-menu-floating");
  menuEl.closest("td")?.classList.add("action-menu-open-cell");

  const rect = toggleBtn.getBoundingClientRect();
  const menuWidth = menuEl.offsetWidth || 170;
  const menuHeight = menuEl.offsetHeight || 120;
  const viewportWidth = window.innerWidth;
  const viewportHeight = window.innerHeight;

  let left = rect.right - menuWidth;
  if (left < 8) left = 8;
  if (left + menuWidth > viewportWidth - 8) left = viewportWidth - menuWidth - 8;

  let top = rect.bottom + 4;
  if (top + menuHeight > viewportHeight - 8) {
    top = rect.top - menuHeight - 4;
  }
  if (top < 8) top = 8;

  menuEl.style.left = `${Math.round(left)}px`;
  menuEl.style.top = `${Math.round(top)}px`;
}

function enableHorizontalDragScroll(container) {
  if (!container) return;
  let isDown = false;
  let startX = 0;
  let startScrollLeft = 0;

  container.addEventListener("mousedown", (event) => {
    const target = event.target;
    if (target instanceof HTMLElement && target.closest("button,select,input,textarea,a,label")) return;
    isDown = true;
    startX = event.pageX;
    startScrollLeft = container.scrollLeft;
  });

  window.addEventListener("mouseup", () => {
    isDown = false;
  });

  container.addEventListener("mouseleave", () => {
    isDown = false;
  });

  container.addEventListener("mousemove", (event) => {
    if (!isDown) return;
    event.preventDefault();
    const delta = event.pageX - startX;
    container.scrollLeft = startScrollLeft - delta;
  });
}

function logActivity(module, action, detail, metadata = {}) {
  const actor = getCurrentUser();
  const entry = {
    id: `a-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
    actor: actor ? actor.username : "system",
    module,
    action,
    detail,
    createdAt: Date.now(),
    ...metadata
  };

  activityLogs.unshift(entry);
  if (activityLogs.length > 300) activityLogs = activityLogs.slice(0, 300);
  saveJSON(STORAGE.activities, activityLogs);
}

function addToRecycleBin(entityType, payload, label) {
  const restoreRef = `r-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
  recycleBin.unshift({
    restoreRef,
    entityType,
    payload,
    label,
    deletedAt: Date.now(),
    deletedBy: getCurrentUser()?.username || "system",
    restoredAt: null
  });
  saveJSON(STORAGE.recycleBin, recycleBin);
  return restoreRef;
}

function restoreDeletedRecord(restoreRef) {
  const item = recycleBin.find((entry) => entry.restoreRef === restoreRef);
  if (!item || item.restoredAt) return false;

  if (item.entityType === "user") {
    const exists = users.some((u) => u.id === item.payload.id);
    if (!exists) {
      const usernameTaken = users.some((u) => u.username.toLowerCase() === item.payload.username.toLowerCase());
      const restoredUser = usernameTaken
        ? { ...item.payload, username: `${item.payload.username}-restored-${Date.now().toString().slice(-4)}` }
        : item.payload;
      users.push(restoredUser);
      saveJSON(STORAGE.users, users);
      syncUsersToRemoteInBackground(`Khôi phục tài khoản ${restoredUser.username}`);
    }
  } else if (item.entityType === "customer") {
    const exists = customers.some((c) => c.id === item.payload.id);
    if (!exists) {
      customers.unshift(item.payload);
      saveJSON(STORAGE.customers, customers);
    }
  } else {
    return false;
  }

  item.restoredAt = Date.now();
  saveJSON(STORAGE.recycleBin, recycleBin);
  logActivity("Khôi phục", "Khôi phục dữ liệu", item.label, { restoreRef, restoredAt: item.restoredAt });
  return true;
}

function restoreActivityAction(activityId) {
  const item = activityLogs.find((entry) => entry.id === activityId);
  if (!item || item.restoredAt || !item.restoreAction) return false;

  const action = item.restoreAction;
  if (action.kind === "news-post-edit" && action.previousPost) {
    const prev = clonePlain(action.previousPost);
    const idx = newsPosts.findIndex((entry) => entry.id === prev.id);
    if (idx !== -1) newsPosts[idx] = prev;
    else newsPosts.unshift(prev);
    saveJSON(STORAGE.newsPosts, newsPosts);
  } else if (action.kind === "news-post-delete" && action.deletedPost) {
    const post = clonePlain(action.deletedPost);
    const exists = newsPosts.some((entry) => entry.id === post.id);
    if (!exists) newsPosts.unshift(post);
    saveJSON(STORAGE.newsPosts, newsPosts);
  } else if (action.kind === "news-event-edit" && action.previousEvent) {
    const prev = clonePlain(action.previousEvent);
    const idx = newsEvents.findIndex((entry) => entry.id === prev.id);
    if (idx !== -1) newsEvents[idx] = prev;
    else newsEvents.unshift(prev);
    saveJSON(STORAGE.newsEvents, newsEvents);
  } else if (action.kind === "news-event-delete" && action.deletedEvent) {
    const eventData = clonePlain(action.deletedEvent);
    const exists = newsEvents.some((entry) => entry.id === eventData.id);
    if (!exists) newsEvents.unshift(eventData);
    saveJSON(STORAGE.newsEvents, newsEvents);
  } else if (action.kind === "schedule-edit" && action.previousSchedule) {
    const prev = clonePlain(action.previousSchedule);
    const idx = schedules.findIndex((entry) => entry.id === prev.id);
    if (idx !== -1) schedules[idx] = prev;
    else schedules.unshift(prev);
    saveJSON(STORAGE.schedule, schedules);
  } else if (action.kind === "schedule-delete" && action.deletedSchedule) {
    const scheduleData = clonePlain(action.deletedSchedule);
    const exists = schedules.some((entry) => entry.id === scheduleData.id);
    if (!exists) schedules.unshift(scheduleData);
    saveJSON(STORAGE.schedule, schedules);
  } else {
    return false;
  }

  item.restoredAt = Date.now();
  saveJSON(STORAGE.activities, activityLogs);
  logActivity("Khôi phục", "Khôi phục chỉnh sửa", item.detail || item.action, { restoredAt: item.restoredAt });
  return true;
}

function getActivityFilteredRows() {
  let start = activityViewState.start || "";
  let end = activityViewState.end || "";
  if (!start && !end) return activityLogs;
  if (!start) start = end;
  if (!end) end = start;
  if (start > end) {
    const temp = start;
    start = end;
    end = temp;
  }
  return activityLogs.filter((entry) => {
    const key = toDateKey(new Date(entry.createdAt));
    return key >= start && key <= end;
  });
}

function renderActivityTable() {
  const isAdmin = can("canManageUsers");
  els.activityHint.textContent = isAdmin
    ? "Lịch sử mới nhất của các thao tác chỉnh sửa trên hệ thống."
    : "Chỉ Admin được xem lịch sử chỉnh sửa và thao tác trên hệ thống.";

  if (!isAdmin) {
    els.activityBody.innerHTML = "<tr><td colspan=\"6\">Bạn không có quyền xem lịch sử hoạt động.</td></tr>";
    if (els.activityPageInfo) els.activityPageInfo.textContent = "Trang 0/0";
    if (els.activityPrevPageBtn) els.activityPrevPageBtn.disabled = true;
    if (els.activityNextPageBtn) els.activityNextPageBtn.disabled = true;
    return;
  }

  if (els.activityFilterStartDate) els.activityFilterStartDate.value = activityViewState.start || "";
  if (els.activityFilterEndDate) els.activityFilterEndDate.value = activityViewState.end || "";

  const rows = getActivityFilteredRows();
  const totalPages = Math.max(1, Math.ceil(rows.length / activityViewState.pageSize));
  if (activityViewState.page > totalPages) activityViewState.page = totalPages;
  if (activityViewState.page < 1) activityViewState.page = 1;

  const start = (activityViewState.page - 1) * activityViewState.pageSize;
  const paginatedRows = rows.slice(start, start + activityViewState.pageSize);

  if (els.activityPageInfo) {
    els.activityPageInfo.textContent = `Trang ${activityViewState.page}/${totalPages} • ${rows.length} bản ghi`;
  }
  if (els.activityPrevPageBtn) els.activityPrevPageBtn.disabled = activityViewState.page <= 1;
  if (els.activityNextPageBtn) els.activityNextPageBtn.disabled = activityViewState.page >= totalPages;

  if (rows.length === 0) {
    els.activityBody.innerHTML = "<tr><td colspan=\"6\">Chưa có lịch sử hoạt động.</td></tr>";
    return;
  }

  els.activityBody.innerHTML = paginatedRows
    .map((item) => `
      <tr>
        <td>${new Date(item.createdAt).toLocaleString("vi-VN", { hour12: false })}</td>
        <td>${item.actor}</td>
        <td>${item.module}</td>
        <td>${item.action}</td>
        <td>${item.detail}</td>
        <td>
          ${(item.restoreRef || item.restoreAction) && !item.restoredAt
            ? `<button class="btn secondary" type="button" data-restore-ref="${item.restoreRef || ""}" data-restore-activity-id="${item.id}">Khôi phục</button>`
            : item.restoredAt
              ? "Đã khôi phục"
              : "--"}
        </td>
      </tr>
    `)
    .join("");
}

function setBrandLogo(src) {
  if (!src) {
    els.brandTitleLogo.removeAttribute("src");
    els.brandTitleLogo.classList.add("hidden");
    els.brandTitleText.classList.remove("hidden");
    return;
  }
  els.brandTitleLogo.src = src;
  els.brandTitleLogo.classList.remove("hidden");
  els.brandTitleText.classList.add("hidden");
}

async function detectLogoFromPublic() {
  for (const path of LOGO_CANDIDATES) {
    try {
      const response = await fetch(path, { method: "HEAD", cache: "no-store" });
      if (response.ok) return path;
    } catch {
      // Continue trying next candidate.
    }
  }
  return null;
}

async function initBrandLogo() {
  const savedLogo = localStorage.getItem(STORAGE.logo);
  if (savedLogo) {
    setBrandLogo(savedLogo);
    return;
  }

  const detected = await detectLogoFromPublic();
  if (detected) setBrandLogo(detected);
}

function getTodayLatestRows(list) {
  const map = new Map();
  list.filter((r) => r.date === today).forEach((row) => {
    const prev = map.get(row.department);
    if (!prev || prev.updatedAt < row.updatedAt) map.set(row.department, row);
  });
  return Array.from(map.values());
}

function toDateKey(dateObj) {
  return new Date(dateObj.getTime() - dateObj.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
}

function getPresetRange(preset) {
  const now = new Date();
  const end = toDateKey(now);
  const startDate = new Date(now);

  if (preset === "7d") {
    startDate.setDate(startDate.getDate() - 6);
  } else if (preset === "30d") {
    startDate.setDate(startDate.getDate() - 29);
  } else if (preset === "thisMonth") {
    startDate.setDate(1);
  }

  const start = preset === "today" ? end : toDateKey(startDate);
  return { start, end };
}

function setFilterInputs() {
  els.timePreset.value = filterState.preset;
  els.filterStart.value = filterState.start;
  els.filterEnd.value = filterState.end;
}

function setFilterSummary() {
  if (filterState.start === filterState.end) {
    els.filterSummary.textContent = `Bộ lọc hiện tại: ${filterState.start}`;
    return;
  }
  els.filterSummary.textContent = `Bộ lọc hiện tại: ${filterState.start} đến ${filterState.end}`;
}

function applyPresetFilter(preset) {
  const range = getPresetRange(preset);
  filterState = { preset, start: range.start, end: range.end };
  setFilterInputs();
  setFilterSummary();
}

function applyCustomRange() {
  let start = els.filterStart.value || filterState.start;
  let end = els.filterEnd.value || filterState.end;
  if (start > end) {
    const temp = start;
    start = end;
    end = temp;
  }
  filterState = { preset: "custom", start, end };
  setFilterInputs();
  setFilterSummary();
}

function getFilteredReports() {
  const start = filterState.start;
  const end = filterState.end;
  return reports.filter((row) => row.date >= start && row.date <= end);
}

function getRolePermissions(roleKey) {
  return rolePermissionsState[roleKey] || rolePermissionsState.staff || ROLES.staff;
}

function getCurrentUser() {
  if (!authState.loggedIn || !authState.userId) return null;
  return users.find((u) => u.id === authState.userId) || null;
}

function can(permission) {
  const user = getCurrentUser();
  if (!user) return false;
  const perms = getRolePermissions(user.roleKey);
  return Boolean(perms[permission]);
}

function closeMenu() {
  els.appMenu.classList.remove("open");
  els.menuOverlay.classList.add("hidden");
}

function openMenu() {
  if (!authState.loggedIn) return;
  els.appMenu.classList.add("open");
  els.menuOverlay.classList.remove("hidden");
}

function updateBackButtonState() {
  if (!els.backBtn) return;
  const disabled = !authState.loggedIn || pageHistory.length === 0;
  els.backBtn.disabled = disabled;
}

function closeCustomerModal() {
  els.customerModal.classList.add("hidden");
}

function resetCustomerForm() {
  editingCustomerId = null;
  els.customerModalTitle.textContent = "Thêm khách hàng mới";
  els.saveCustomerBtn.textContent = "Lưu khách hàng";
  renderCustomerOwnerOptions(getCurrentUser()?.username || "Chưa gán");
  els.customerName.value = "";
  els.customerContactPerson.value = "";
  els.customerPhone.value = "";
  els.customerEmail.value = "";
  els.customerAddress.value = "";
  els.customerAddrProvince.value = "";
  els.customerAddrDistrict.innerHTML = '<option value="">Lựa chọn</option>';
  els.customerAddrDistrict.disabled = true;
  els.customerAddrWard.innerHTML = '<option value="">Lựa chọn</option>';
  els.customerAddrWard.disabled = true;
  els.customerTier.value = "Tiêu chuẩn";
  els.customerStatus.value = "Đã gọi";
  els.customerOwner.value = getCurrentUser()?.username || "Chưa gán";
  els.customerSource.value = "Nhập tay";
  els.customerDemand.value = "";
  els.customerNote.value = "";
}

function openCustomerModal() {
  if (!editingCustomerId) resetCustomerForm();
  els.customerModal.classList.remove("hidden");
}

function setActivePage(pageKey, options = {}) {
  const { fromHistory = false, syncUrl = true, replaceUrl = false } = options;
  if (!canAccessPage(pageKey)) {
    pageKey = "home";
  }

  if (pageKey !== "workflow") {
    activeWorkflowDetail = null;
  }

  if (pageKey === "schedule") {
    scheduleFilterState = { month: today.slice(0, 7), status: "", staff: "all", source: "", keyword: "" };
    renderScheduleStaffControls();
    if (els.scheduleFilterMonth) els.scheduleFilterMonth.value = scheduleFilterState.month;
    if (els.scheduleFilterStatus) els.scheduleFilterStatus.value = "";
    if (els.scheduleFilterStaff) els.scheduleFilterStaff.value = "all";
    if (els.scheduleFilterSource) els.scheduleFilterSource.value = "";
    if (els.scheduleSearch) els.scheduleSearch.value = "";
    renderScheduleTable();
    setTimeout(() => {
      els.scheduleTableWrap?.focus();
    }, 0);
  }

  if (!fromHistory && authState.loggedIn && activePage !== pageKey) {
    pageHistory.push(activePage);
    if (pageHistory.length > 20) pageHistory.shift();
  }

  activePage = pageKey;
  els.appPages.forEach((page) => {
    page.classList.toggle("hidden", page.dataset.page !== pageKey);
  });
  els.menuItems.forEach((item) => {
    item.classList.toggle("active", item.dataset.page === pageKey);
  });
  if (authState.loggedIn && syncUrl && !fromHistory) {
    syncUrlWithPage(pageKey, { replace: replaceUrl });
  }
  updateBackButtonState();
  closeMenu();
}

function renderWorkflowDetailView() {
  if (!els.workflowSystemView || !els.workflowDetailView || !els.workflowDetailTitle || !els.workflowDetailSteps) return;

  const detail = activeWorkflowDetail ? WORKFLOW_DEPARTMENT_DETAILS[activeWorkflowDetail] : null;
  const isDetailOpen = Boolean(detail);

  els.workflowSystemView.classList.toggle("hidden", isDetailOpen);
  els.workflowDetailView.classList.toggle("hidden", !isDetailOpen);

  if (!detail) return;

  els.workflowDetailTitle.textContent = `Quy trinh chi tiet - ${detail.title}`;
  els.workflowDetailSteps.innerHTML = detail.steps.map((step) => `<li>${step}</li>`).join("");
}

function openWorkflowDepartmentDetail(departmentKey) {
  if (!WORKFLOW_DEPARTMENT_DETAILS[departmentKey]) return;
  activeWorkflowDetail = departmentKey;
  renderWorkflowDetailView();
}

function closeWorkflowDepartmentDetail() {
  activeWorkflowDetail = null;
  renderWorkflowDetailView();
}

function renderPolicyFolders() {
  if (!els.policyFolderRoot) return;

  const parentEntries = Object.entries(POLICY_FOLDER_STRUCTURE);
  els.policyFolderRoot.innerHTML = parentEntries
    .map(([parentKey, parentValue]) => {
      const isParentOpen = activePolicyParent === parentKey;
      const parentArrow = isParentOpen ? "▾" : "▸";
      const childrenHtml = Object.entries(parentValue.children)
        .map(([childKey, childRules]) => {
          const childId = `${parentKey}__${childKey}`;
          const isChildOpen = activePolicyChild === childId;
          const childArrow = isChildOpen ? "▾" : "▸";
          const rulesHtml = childRules.map((rule) => `<li>${rule}</li>`).join("");
          return `
            <div class="card" style="padding:10px;border:1px solid #e2e8f0;">
              <button
                class="btn secondary"
                type="button"
                data-policy-child="${childId}"
                style="width:100%;display:flex;justify-content:space-between;align-items:center;background:#f8fafc;color:#0f172a;"
              >
                <span>📂 ${POLICY_DEPARTMENT_LABELS[childKey] || childKey}</span>
                <span>${childArrow}</span>
              </button>
              <div class="${isChildOpen ? "" : "hidden"}" style="margin-top:8px;padding:10px;background:#ffffff;border:1px dashed #cbd5e1;border-radius:8px;">
                <ul style="margin:0;padding-left:20px;line-height:1.8;">${rulesHtml}</ul>
              </div>
            </div>
          `;
        })
        .join("");

      return `
        <div class="card" style="margin-bottom:10px;padding:10px;border-left:4px solid ${parentValue.color};">
          <button
            class="btn"
            type="button"
            data-policy-parent="${parentKey}"
            style="width:100%;display:flex;justify-content:space-between;align-items:center;background:${parentValue.color};color:#fff;"
          >
            <span>📁 ${parentValue.title}</span>
            <span>${parentArrow}</span>
          </button>
          <div class="${isParentOpen ? "" : "hidden"}" style="margin-top:10px;display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;">
            ${childrenHtml}
          </div>
        </div>
      `;
    })
    .join("");
}

function togglePolicyParent(parentKey) {
  if (!POLICY_FOLDER_STRUCTURE[parentKey]) return;
  if (activePolicyParent === parentKey) {
    activePolicyParent = null;
    activePolicyChild = null;
  } else {
    activePolicyParent = parentKey;
    activePolicyChild = null;
  }
  renderPolicyFolders();
}

function togglePolicyChild(childId) {
  const [parentKey, childKey] = (childId || "").split("__");
  if (!parentKey || !childKey || !POLICY_FOLDER_STRUCTURE[parentKey]?.children[childKey]) return;
  activePolicyParent = parentKey;
  activePolicyChild = activePolicyChild === childId ? null : childId;
  renderPolicyFolders();
}

function performLogout() {
  authState = { loggedIn: false, role: null, username: null, userId: null };
  saveJSON(STORAGE.auth, authState);
  activePage = "home";
  pageHistory = [];
  renderAll();
}

function enforceActiveSessionAccess() {
  if (!authState.loggedIn || !authState.userId) return false;
  const currentUser = users.find((u) => u.id === authState.userId);
  if (!currentUser) {
    showToast("Tài khoản đã bị xóa. Vui lòng liên hệ quản trị viên.", "warning");
    performLogout();
    return true;
  }
  if ((currentUser.status || "active") === "suspended") {
    showToast("Tài khoản đang bị tạm dừng. Bạn đã được đăng xuất.", "warning");
    performLogout();
    return true;
  }
  return false;
}

function setAuthUI() {
  if (authState.loggedIn) {
    const user = getCurrentUser();
    if (!user || (user.status || "active") === "suspended") {
      authState = { loggedIn: false, role: null, username: null, userId: null };
      saveJSON(STORAGE.auth, authState);
      if (user && (user.status || "active") === "suspended") {
        showToast("Tài khoản đang bị tạm dừng. Bạn đã được đăng xuất.", "warning");
      }
      setAuthUI();
      return;
    }

    const roleLabel = getRolePermissions(user.roleKey).label;
    els.authMessage.textContent = `Đăng nhập: ${user.fullName} (${roleLabel})`;
    els.logoutBtn.classList.remove("hidden");
    els.loginBtn.classList.add("hidden");
    els.loginPassword.value = "";
    els.loginUsername.value = user.username;
    els.loginUsername.disabled = true;
    els.postLoginOnly.forEach((section) => section.classList.remove("hidden"));
    els.headerActions.classList.remove("hidden");
    els.brandBlock.classList.add("logo-editable");
    els.dashboardRoot.classList.remove("logged-out");
    els.dashboardRoot.classList.add("logged-in");
    setActivePage(activePage);
  } else {
    els.authMessage.textContent = "Chưa đăng nhập";
    els.logoutBtn.classList.add("hidden");
    els.loginBtn.classList.remove("hidden");
    els.loginUsername.disabled = false;
    applyLoginPrefsToForm();
    els.postLoginOnly.forEach((section) => section.classList.add("hidden"));
    els.headerActions.classList.add("hidden");
    els.brandBlock.classList.remove("logo-editable");
    els.dashboardRoot.classList.add("logged-out");
    els.dashboardRoot.classList.remove("logged-in");
    pageHistory = [];
    updateBackButtonState();
    closeMenu();
  }

  const canViewData = can("canViewData");
  const canViewUsers = can("canViewUsers");
  const isAdmin = can("canManageUsers");
  const canManageInventory = can("canSubmitReport");
  els.syncBtn.disabled = !can("canSyncData");
  els.testConnectionBtn.disabled = !can("canSyncData");
  els.pdfBtn.disabled = !can("canExportPdf");
  if (els.syncTelegramBtn) els.syncTelegramBtn.disabled = !can("canSyncData");
  if (els.testTelegramBtn) els.testTelegramBtn.disabled = !can("canSyncData");
  const telegramSection = document.getElementById("telegramSourceSection");
  if (telegramSection) telegramSection.style.display = isAdmin ? "" : "none";
  els.submitReportBtn.disabled = !can("canSubmitReport");
  if (els.openInventoryModalBtn) els.openInventoryModalBtn.disabled = !canManageInventory;
  if (els.exportInventoryExcelBtn) els.exportInventoryExcelBtn.disabled = !canManageInventory;
  if (els.exportInventoryPdfBtn) els.exportInventoryPdfBtn.disabled = !can("canExportPdf");
  if (els.exportReportsExcelBtn) els.exportReportsExcelBtn.disabled = !canViewData;
  if (els.exportReportsPdfBtn) els.exportReportsPdfBtn.disabled = !can("canExportPdf");
  if (els.exportCareExcelBtn) els.exportCareExcelBtn.disabled = !canViewData;
  if (els.exportCarePdfBtn) els.exportCarePdfBtn.disabled = !can("canExportPdf");
  if (els.applyInventoryStatsBtn) els.applyInventoryStatsBtn.disabled = !canManageInventory;
  if (els.resetInventoryStatsBtn) els.resetInventoryStatsBtn.disabled = !canManageInventory;
  if (els.openUserModalBtn) els.openUserModalBtn.disabled = !isAdmin;
  if (els.openPermissionModalBtn) els.openPermissionModalBtn.disabled = !isAdmin;
  els.adminUserSection.classList.toggle("hidden", activePage !== "hr" || !canViewUsers);
  els.activitySection.classList.toggle("hidden", activePage !== "activity" || !isAdmin);
  els.menuItems.forEach((item) => {
    const pageKey = item.dataset.page;
    if (!pageKey) return;
    item.classList.toggle("hidden", !canAccessPage(pageKey));
  });

  if (!canViewData) {
    els.kpiReports.textContent = "0";
    els.kpiCompletion.textContent = "0%";
    els.kpiQuality.textContent = "0";
    els.kpiIssues.textContent = "0";
    els.reportBody.innerHTML = "";
  }
}

function getFilteredUsers() {
  const keyword = (els.hrQuickSearch?.value || "").toLowerCase().trim();
  const role = els.hrFilterRole?.value || "";
  const dept = els.hrFilterDept?.value || "";
  const status = els.hrFilterStatus?.value || "";
  return users.filter((u) => {
    if (keyword && !u.fullName.toLowerCase().includes(keyword) && !u.username.toLowerCase().includes(keyword) && !(u.userCode || "").toLowerCase().includes(keyword)) return false;
    if (role && u.roleKey !== role) return false;
    if (dept && (u.department || "Ban điều hành") !== dept) return false;
    if (status && (u.status || "active") !== status) return false;
    return true;
  });
}

function renderUserTable() {
  const filtered = getFilteredUsers();
  const canManageUsers = can("canManageUsers");
  if (filtered.length === 0) {
    els.userBody.innerHTML = '<tr><td colspan="9" style="text-align:center;">Không tìm thấy nhân sự phù hợp.</td></tr>';
    return;
  }
  els.userBody.innerHTML = filtered
    .sort((a, b) => (a.userCode || "").localeCompare(b.userCode || "") || a.username.localeCompare(b.username))
    .map((u, idx) => {
      const roleLabel = getRolePermissions(u.roleKey).label;
      const createdAt = u.createdAt ? new Date(u.createdAt).toLocaleDateString("vi-VN") : "--";
      const isSuspended = (u.status || "active") === "suspended";
      const statusBadge = isSuspended
        ? '<span style="background:#fde8e8;color:#c0392b;border-radius:12px;padding:2px 10px;font-size:0.78rem;font-weight:600;">Tạm dừng</span>'
        : '<span style="background:#e8f8ef;color:#1a7f4b;border-radius:12px;padding:2px 10px;font-size:0.78rem;font-weight:600;">Hoạt động</span>';
      return `
      <tr${isSuspended ? ' style="opacity:0.6;"' : ''}>
        <td class="muted">${idx + 1}</td>
        <td style="font-weight:600;letter-spacing:0.5px;"><span style="background:#eef2ff;color:#4338ca;border-radius:8px;padding:2px 8px;font-size:0.82rem;">${u.userCode || "--"}</span></td>
        <td><button class="user-name-link" data-user-id="${u.id}" type="button">${u.fullName}</button></td>
        <td class="muted">${u.username}</td>
        <td>${roleLabel}</td>
        <td>${u.department || "Ban điều hành"}</td>
        <td>${statusBadge}</td>
        <td class="muted" style="font-size:0.8rem;">${createdAt}</td>
        <td style="position:relative;">
          ${canManageUsers ? `
          <button class="btn secondary user-action-toggle" data-user-id="${u.id}" type="button" title="Thao tác">⋯</button>
          <div class="user-action-menu hidden" data-user-id="${u.id}">
            <button class="user-action-item user-edit-btn" data-user-id="${u.id}" type="button">✏️ Sửa</button>
            <button class="user-action-item user-suspend-btn" data-user-id="${u.id}" type="button">${isSuspended ? '▶️ Kích hoạt' : '⏸ Tạm dừng'}</button>
            <button class="user-action-item user-action-item--danger user-delete-btn" data-user-id="${u.id}" type="button">🗑 Xóa</button>
          </div>
          ` : '<span class="muted" style="font-size:0.8rem;">Chỉ xem</span>'}
        </td>
      </tr>
      `;
    })
    .join("");
}

function formatCurrency(value) {
  return `${Number(value || 0).toLocaleString("vi-VN")} đ`;
}

function getInventoryStockLevel(item) {
  const qty = Number(item.quantity) || 0;
  const threshold = Number(item.alertThreshold) || 0;
  if (qty <= 0) return "out";
  if (qty <= threshold) return "low";
  return "safe";
}

function getInventoryStockLabel(level) {
  if (level === "out") return "Hết hàng";
  if (level === "low") return "Gần hết";
  return "An toàn";
}

function getInventoryStatusLabel(status) {
  return status === "inactive" ? "Ngừng kinh doanh" : "Đang kinh doanh";
}

function getFilteredInventoryItems() {
  const keyword = (els.inventorySearch.value || "").toLowerCase().trim();
  const status = els.inventoryFilterStatus.value || "";
  const stock = els.inventoryFilterStock.value || "";

  return inventoryItems.filter((item) => {
    const haystack = `${item.productCode || ""} ${item.productName || ""}`.toLowerCase();
    if (keyword && !haystack.includes(keyword)) return false;
    if (status && item.status !== status) return false;
    const stockLevel = getInventoryStockLevel(item);
    if (stock && stockLevel !== stock) return false;
    return true;
  });
}

function isDateInRange(dateKey, start, end) {
  if (start && dateKey < start) return false;
  if (end && dateKey > end) return false;
  return true;
}

function getFilteredInventoryTransactions() {
  const productKeyword = (els.inventorySearch.value || "").toLowerCase().trim();
  const status = els.inventoryFilterStatus.value || "";
  const stock = els.inventoryFilterStock.value || "";
  const { start, end } = inventoryStatsState;

  return inventoryTransactions.filter((txn) => {
    const item = inventoryItems.find((row) => row.id === txn.itemId);
    const itemStatus = item?.status || "active";
    const itemStock = item ? getInventoryStockLevel(item) : "safe";
    const haystack = `${txn.productCode || ""} ${txn.productName || ""}`.toLowerCase();
    const txnDate = new Date(txn.createdAt).toISOString().slice(0, 10);

    if (productKeyword && !haystack.includes(productKeyword)) return false;
    if (status && itemStatus !== status) return false;
    if (stock && itemStock !== stock) return false;
    if (!isDateInRange(txnDate, start, end)) return false;
    return true;
  });
}

function renderInventoryStatsAndHistory() {
  const txns = getFilteredInventoryTransactions().sort((a, b) => b.createdAt - a.createdAt);
  const totalIn = txns.filter((t) => t.type === "in").reduce((sum, t) => sum + t.quantity, 0);
  const totalOut = txns.filter((t) => t.type === "out").reduce((sum, t) => sum + t.quantity, 0);
  const net = totalIn - totalOut;

  els.inventoryInTotal.textContent = totalIn.toLocaleString("vi-VN");
  els.inventoryOutTotal.textContent = totalOut.toLocaleString("vi-VN");
  els.inventoryTxnTotal.textContent = txns.length.toLocaleString("vi-VN");
  els.inventoryNetTotal.textContent = net.toLocaleString("vi-VN");
  els.inventoryStatsRangeText.textContent = inventoryStatsState.start || inventoryStatsState.end
    ? `Khoảng: ${inventoryStatsState.start || "--"} -> ${inventoryStatsState.end || "--"}`
    : "Khoảng: Toàn thời gian";

  if (!txns.length) {
    els.inventoryHistoryBody.innerHTML = '<tr><td colspan="8" style="text-align:center;">Chưa có lịch sử xuất/nhập trong khoảng thời gian đã chọn.</td></tr>';
    return;
  }

  els.inventoryHistoryBody.innerHTML = txns
    .map((txn) => {
      const typeLabel = txn.type === "in" ? "Nhập" : "Xuất";
      const typeStyle = txn.type === "in"
        ? "background:#ecfdf5;color:#166534;"
        : "background:#fff1f2;color:#be123c;";
      return `
      <tr>
        <td class="muted" style="font-size:0.8rem;">${new Date(txn.createdAt).toLocaleString("vi-VN", { hour12: false })}</td>
        <td style="font-weight:600;">${txn.productCode}</td>
        <td>${txn.productName}</td>
        <td><span style="${typeStyle}border-radius:10px;padding:2px 8px;font-size:0.78rem;font-weight:600;">${typeLabel}</span></td>
        <td><strong>${txn.quantity.toLocaleString("vi-VN")}</strong></td>
        <td>${txn.stockAfter.toLocaleString("vi-VN")}</td>
        <td>${txn.actor || "--"}</td>
        <td class="muted" style="font-size:0.8rem;">${txn.note || "--"}</td>
      </tr>
      `;
    })
    .join("");
}

function renderInventoryTable() {
  const filtered = getFilteredInventoryItems();
  els.inventoryFilterSummary.textContent = `Hiển thị ${filtered.length}/${inventoryItems.length} vật tư`;
  const totalValue = inventoryItems.reduce((sum, item) => sum + (item.purchasePrice || 0) * (item.quantity || 0), 0);
  els.inventoryTotalValue.textContent = formatCurrency(totalValue);

  if (filtered.length === 0) {
    els.inventoryBody.innerHTML = '<tr><td colspan="13" style="text-align:center;">Không có vật tư phù hợp bộ lọc.</td></tr>';
    return;
  }

  els.inventoryBody.innerHTML = filtered
    .sort((a, b) => b.updatedAt - a.updatedAt)
    .map((item, idx) => {
      const stockLevel = getInventoryStockLevel(item);
      const stockStyle = stockLevel === "out"
        ? "background:#fee2e2;color:#b91c1c;"
        : stockLevel === "low"
          ? "background:#fff7ed;color:#c2410c;"
          : "background:#ecfdf5;color:#166534;";
      const productStatusStyle = item.status === "inactive"
        ? "background:#f1f5f9;color:#475569;"
        : "background:#e0f2fe;color:#075985;";
      const lowWarning = stockLevel === "low" || stockLevel === "out"
        ? `<span style="color:#c2410c;font-weight:600;">Cảnh báo: chỉ còn ${item.quantity}</span>`
        : "--";
      return `
      <tr>
        <td class="muted">${idx + 1}</td>
        <td style="font-weight:600;letter-spacing:0.4px;">${item.productCode}</td>
        <td>${item.productName}</td>
        <td>${formatCurrency(item.purchasePrice)}</td>
        <td>${formatCurrency(item.salePrice)}</td>
        <td><strong>${item.quantity}</strong></td>
        <td class="muted" style="font-size:0.85rem;">${item.supplier || "—"}</td>
        <td class="muted" style="font-size:0.85rem;">${(function(){ if (!item.expiryDate) return '—'; const d = new Date(item.expiryDate); const now = new Date(); const expired = d < now; const soon = !expired && (d - Date.now() < 30*24*3600*1000); const color = expired ? '#b91c1c' : (soon ? '#c2410c' : 'inherit'); const fw = (expired || soon) ? '600' : '400'; return '<span style="color:' + color + ';font-weight:' + fw + '">' + item.expiryDate + '</span>'; })()}</td>
        <td style="font-size:0.88rem;font-weight:600;color:#166534;">${formatCurrency(item.purchasePrice * item.quantity)}</td>
        <td><span style="${productStatusStyle}border-radius:10px;padding:2px 8px;font-size:0.78rem;font-weight:600;">${getInventoryStatusLabel(item.status)}</span></td>
        <td>
          <span style="${stockStyle}border-radius:10px;padding:2px 8px;font-size:0.78rem;font-weight:600;">${getInventoryStockLabel(stockLevel)}</span>
          <div class="muted" style="font-size:0.74rem;margin-top:4px;">Ngưỡng: ${item.alertThreshold}</div>
        </td>
        <td class="muted" style="font-size:0.8rem;">${new Date(item.updatedAt).toLocaleString("vi-VN", { hour12: false })}</td>
        <td style="position:relative;">
          <button class="btn secondary user-action-toggle" data-item-id="${item.id}" type="button" title="Thao tác">⋯</button>
          <div class="user-action-menu inventory-action-menu hidden" data-item-id="${item.id}">
            <button class="user-action-item inventory-in-btn" data-item-id="${item.id}" type="button">📥 Nhập kho</button>
            <button class="user-action-item inventory-out-btn" data-item-id="${item.id}" type="button">📤 Xuất kho</button>
            <button class="user-action-item inventory-edit-btn" data-item-id="${item.id}" type="button">✏️ Sửa</button>
            <button class="user-action-item user-action-item--danger inventory-delete-btn" data-item-id="${item.id}" type="button">🗑 Xóa</button>
          </div>
        </td>
      </tr>
      ${stockLevel === "low" || stockLevel === "out" ? `<tr><td colspan="13" style="background:#fffbeb;font-size:0.82rem;">${lowWarning}</td></tr>` : ""}
      `;
    })
    .join("");
}

function resetInventoryForm() {
  editingInventoryId = null;
  els.inventoryModalTitle.textContent = "Thêm vật tư mới";
  els.saveInventoryBtn.textContent = "Lưu vật tư";
  els.inventoryProductCode.value = "";
  els.inventoryProductName.value = "";
  els.inventoryPurchasePrice.value = "";
  els.inventorySupplier.value = "";
  els.inventoryExpiryDate.value = "";
  els.inventorySalePrice.value = "";
  els.inventoryQuantity.value = "";
  els.inventoryAlertThreshold.value = "20";
  els.inventoryProductStatus.value = "active";
  els.inventoryModalStatus.textContent = "";
}

function openInventoryModal() {
  resetInventoryForm();
  els.inventoryModal.classList.remove("hidden");
}

function closeInventoryModal() {
  els.inventoryModal.classList.add("hidden");
  resetInventoryForm();
}

function openInventoryTxnModal(itemId, txnType) {
  const item = inventoryItems.find((i) => i.id === itemId);
  if (!item) return;
  inventoryTxnItemId = itemId;
  els.inventoryTxnTitle.textContent = txnType === "out" ? "Xuất kho vật tư" : "Nhập kho vật tư";
  els.inventoryTxnProduct.value = `${item.productCode} · ${item.productName} (Tồn: ${item.quantity})`;
  els.inventoryTxnType.value = txnType;
  els.inventoryTxnQty.value = "";
  els.inventoryTxnNote.value = "";
  els.inventoryTxnStatus.textContent = "";
  els.inventoryTxnModal.classList.remove("hidden");
}

function closeInventoryTxnModal() {
  els.inventoryTxnModal.classList.add("hidden");
  inventoryTxnItemId = null;
  els.inventoryTxnStatus.textContent = "";
}

function exportFilteredInventoryToExcel() {
  const rows = getFilteredInventoryItems().sort((a, b) => b.updatedAt - a.updatedAt);
  if (rows.length === 0) {
    els.inventoryStatusMessage.textContent = "Không có dữ liệu vật tư để xuất Excel theo bộ lọc hiện tại.";
    return;
  }

  const header = ["Mã sản phẩm", "Tên sản phẩm", "Giá nhập", "Giá bán", "Số lượng", "Nhà cung cấp", "Hạn sử dụng", "Giá trị tồn", "Trạng thái", "Tồn kho", "Ngưỡng cảnh báo", "Cập nhật"];
  const records = rows.map((item) => {
    const stockLevel = getInventoryStockLevel(item);
    return [
      item.productCode,
      item.productName,
      item.purchasePrice,
      item.salePrice,
      item.quantity,
      item.supplier || "",
      item.expiryDate || "",
      item.purchasePrice * item.quantity,
      getInventoryStatusLabel(item.status),
      getInventoryStockLabel(stockLevel),
      item.alertThreshold,
      new Date(item.updatedAt).toLocaleString("vi-VN", { hour12: false })
    ];
  });

  const csv = [header, ...records]
    .map((line) => line.map((cell) => toCsvValue(cell)).join(","))
    .join("\n");

  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8;" });
  downloadBlob(blob, `kho-vat-tu-${today}.csv`);
  showToast(`Đã xuất Excel (CSV) ${rows.length} vật tư.`);
  logActivity("Kho", "Xuất Excel kho", `Số vật tư: ${rows.length}`);
  renderActivityTable();
}

async function exportFilteredInventoryToPdf() {
  if (!can("canExportPdf")) {
    els.inventoryStatusMessage.textContent = "Bạn không có quyền xuất PDF kho.";
    return;
  }

  const rows = getFilteredInventoryItems();
  if (rows.length === 0) {
    els.inventoryStatusMessage.textContent = "Không có dữ liệu vật tư để xuất PDF theo bộ lọc hiện tại.";
    return;
  }

  showToast("Đang tạo PDF kho...", "info");
  const canvas = await html2canvas(els.inventorySection, { scale: 2, useCORS: true });
  const image = canvas.toDataURL("image/png");
  const pdf = new jsPDF("p", "mm", "a4");
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const imageHeight = (canvas.height * pageWidth) / canvas.width;

  let heightLeft = imageHeight;
  let position = 0;

  pdf.addImage(image, "PNG", 0, position, pageWidth, imageHeight);
  heightLeft -= pageHeight;

  while (heightLeft > 0) {
    position -= pageHeight;
    pdf.addPage();
    pdf.addImage(image, "PNG", 0, position, pageWidth, imageHeight);
    heightLeft -= pageHeight;
  }

  pdf.save(`kho-vat-tu-${today}.pdf`);
  showToast(`Đã xuất PDF ${rows.length} vật tư.`);
  logActivity("Kho", "Xuất PDF kho", `Số vật tư: ${rows.length}`);
  renderActivityTable();
}

let viewingHrUserId = null;

function openHrProfileModal(userId) {
  viewingHrUserId = userId;
  const user = users.find((u) => u.id === userId);
  if (!user) return;
  const roleLabel = getRolePermissions(user.roleKey).label;
  els.hrProfileTitle.textContent = user.fullName;
  els.hrProfileSubtitle.textContent = `${user.userCode || ""} · ${user.username}`;
  els.hrProfileInfo.innerHTML = [
    ["Vai trò", roleLabel],
    ["Phòng ban", user.department || "Ban điều hành"],
    ["Trạng thái", (user.status || "active") === "suspended" ? "Tạm dừng" : "Hoạt động"],
    ["SĐT", user.phone || "—"],
    ["Email", user.email || "—"],
    ["Địa chỉ", user.address || "—"],
    ["STK ngân hàng", user.bankAccount || "—"],
  ].map(([k, v]) => `<div><div class="muted" style="font-size:0.75rem;margin-bottom:2px;">${k}</div><strong style="word-break:break-word;">${v}</strong></div>`).join("");
  els.hrProfileStatus.textContent = "";
  renderHrFileList(userId);
  els.hrProfileModal.classList.remove("hidden");
}

function closeHrProfileModal() {
  els.hrProfileModal.classList.add("hidden");
  viewingHrUserId = null;
}

function renderHrFileList(userId) {
  const files = hrFiles[userId] || [];
  if (files.length === 0) {
    els.hrProfileFileList.innerHTML = '<p class="muted" style="font-size:0.85rem;text-align:center;padding:16px 0;">Chưa có tài liệu nào. Nhấn "+ Tải lên" để thêm.</p>';
    return;
  }
  const typeIcon = (name) => {
    const ext = (name.split(".").pop() || "").toLowerCase();
    if (["jpg", "jpeg", "png", "gif", "webp"].includes(ext)) return "🖼";
    if (ext === "pdf") return "📄";
    if (["doc", "docx"].includes(ext)) return "📝";
    if (["xls", "xlsx"].includes(ext)) return "📊";
    return "📎";
  };
  els.hrProfileFileList.innerHTML = files.map((f) => `
    <div style="display:flex;align-items:center;gap:10px;padding:8px 12px;background:#f8fafc;border-radius:8px;border:1px solid #e2e8f0;">
      <span style="font-size:1.3rem;">${typeIcon(f.name)}</span>
      <div style="flex:1;min-width:0;">
        <div style="font-size:0.88rem;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${f.name}</div>
        <div class="muted" style="font-size:0.75rem;">${f.fileType || ""} · ${new Date(f.uploadedAt).toLocaleDateString("vi-VN")}</div>
      </div>
      <button class="btn secondary hr-file-download-btn" data-file-id="${f.id}" data-user-id="${userId}" type="button" style="font-size:0.75rem;padding:4px 10px;">Tải xuống</button>
      <button class="btn warn hr-file-delete-btn" data-file-id="${f.id}" data-user-id="${userId}" type="button" style="font-size:0.75rem;padding:4px 10px;">Xóa</button>
    </div>
  `).join("");
}

function generateUserCode() {
  const max = users.reduce((acc, u) => {
    const m = (u.userCode || "").match(/^NR(\d+)$/);
    return m ? Math.max(acc, parseInt(m[1], 10)) : acc;
  }, 0);
  return `NR${String(max + 1).padStart(3, "0")}`;
}

function resetUserForm() {
  editingUserId = null;
  els.userModalTitle.textContent = "Thêm nhân sự mới";
  els.saveUserBtn.textContent = "Tạo tài khoản";
  els.userCode.value = generateUserCode();
  els.userFullName.value = "";
  els.userUsername.value = "";
  els.userPassword.value = "";
  els.userRoleKey.value = "staff";
  els.userDepartment.value = "Ban điều hành";
  els.userPhone.value = "";
  els.userEmail.value = "";
  els.userAddress.value = "";
  els.userBankAccount.value = "";
  els.userManageStatus.textContent = "";
  els.userModalStatus.textContent = "";
}

function openUserModal() {
  resetUserForm();
  els.userModal.classList.remove("hidden");
}

function closeUserModal() {
  els.userModal.classList.add("hidden");
  resetUserForm();
}

function renderCustomerTable() {
  renderCustomerFilterControls();
  const filteredCustomers = getFilteredCustomers();
  setCustomerFilterSummary(filteredCustomers.length, customers.length);

  if (filteredCustomers.length === 0) {
    els.customerBody.innerHTML = "<tr><td colspan=\"9\">Không có khách hàng phù hợp với bộ lọc hiện tại.</td></tr>";
    return;
  }

  els.customerBody.innerHTML = [...filteredCustomers]
    .sort((a, b) => b.updatedAt - a.updatedAt)
    .map((c) => `
      <tr>
        <td>
          <strong>${c.name}</strong>
          <div class="muted" style="font-size:0.78rem;">${c.address || "--"}</div>
        </td>
        <td>
          <div>${c.contactPerson || "--"}</div>
          <div class="muted" style="font-size:0.78rem;">${c.phone || "--"} | ${c.email || "--"}</div>
        </td>
        <td>${c.tier}</td>
        <td>${c.owner || "--"}</td>
        <td>${c.status}</td>
        <td>${c.source || "--"}</td>
        <td>
          <div class="muted" style="font-size:0.78rem;">Nhu cầu: ${c.demand || "--"}</div>
          <div>${c.note || "--"}</div>
        </td>
        <td>${new Date(c.updatedAt).toLocaleString("vi-VN", { hour12: false })}</td>
        <td>
          <button class="btn secondary customer-edit-btn" type="button" data-customer-id="${c.id}">Sửa</button>
          <button class="btn warn customer-delete-btn" type="button" data-customer-id="${c.id}">Xóa</button>
        </td>
      </tr>
    `)
    .join("");
}

function getCustomerDateKey(customer) {
  return toDateKey(new Date(Number(customer.updatedAt) || Date.now()));
}

function getCustomerOptionSets() {
  const ownerSet = new Set(customers.map((item) => item.owner).filter(Boolean));
  users.forEach((user) => ownerSet.add(user.username));
  ownerSet.add("Chưa gán");

  const statusSet = new Set(CUSTOMER_STATUSES);
  customers.forEach((item) => {
    if (item.status) statusSet.add(item.status);
  });

  const sourceSet = new Set(CUSTOMER_SOURCES);
  customers.forEach((item) => {
    if (item.source) sourceSet.add(item.source);
  });

  return {
    owners: Array.from(ownerSet).sort((a, b) => a.localeCompare(b, "vi")),
    statuses: Array.from(statusSet),
    sources: Array.from(sourceSet)
  };
}

function renderCustomerOwnerOptions(selectedOwner = "") {
  const { owners } = getCustomerOptionSets();
  els.customerOwner.innerHTML = owners
    .map((owner) => `<option value="${owner}">${owner}</option>`)
    .join("");

  if (selectedOwner && owners.includes(selectedOwner)) {
    els.customerOwner.value = selectedOwner;
    return;
  }

  const defaultOwner = getCurrentUser()?.username || owners[0] || "Chưa gán";
  if (owners.includes(defaultOwner)) {
    els.customerOwner.value = defaultOwner;
  }
}

function renderCustomerFilterControls() {
  const { owners, statuses, sources } = getCustomerOptionSets();
  const selectedOwner = customerFilterState.owner;
  const selectedStatus = customerFilterState.status;
  const selectedSource = customerFilterState.source;

  els.customerFilterOwner.innerHTML = ["<option value=\"all\">Tất cả người phụ trách</option>"]
    .concat(owners.map((owner) => `<option value="${owner}">${owner}</option>`))
    .join("");
  els.customerFilterStatus.innerHTML = ["<option value=\"all\">Tất cả trạng thái</option>"]
    .concat(statuses.map((status) => `<option value="${status}">${status}</option>`))
    .join("");
  els.customerFilterSource.innerHTML = ["<option value=\"all\">Tất cả nguồn data</option>"]
    .concat(sources.map((source) => `<option value="${source}">${source}</option>`))
    .join("");

  els.customerFilterStartDate.value = customerFilterState.start;
  els.customerFilterEndDate.value = customerFilterState.end;
  updateCustomerDateRangeTrigger();
  els.customerQuickSearch.value = customerFilterState.keyword || "";
  els.customerFilterOwner.value = owners.includes(selectedOwner) ? selectedOwner : "all";
  els.customerFilterStatus.value = statuses.includes(selectedStatus) ? selectedStatus : "all";
  els.customerFilterSource.value = sources.includes(selectedSource) ? selectedSource : "all";
}

function updateCustomerDateRangeTrigger() {
  const start = customerFilterState.start;
  const end = customerFilterState.end;
  if (!start && !end) {
    els.customerDateRangeTrigger.textContent = "Mọi ngày";
    return;
  }
  if (start && end) {
    els.customerDateRangeTrigger.textContent = `${start} - ${end}`;
    return;
  }
  const single = start || end;
  els.customerDateRangeTrigger.textContent = `Ngày: ${single}`;
}

function closeCustomerDateRangePopover() {
  els.customerDateRangePopover.classList.add("hidden");
  els.customerDateRangeTrigger.setAttribute("aria-expanded", "false");
}

function openCustomerDateRangePopover() {
  els.customerDateRangePopover.classList.remove("hidden");
  els.customerDateRangeTrigger.setAttribute("aria-expanded", "true");
}

function applyCustomerDateRangeFromInputs() {
  let start = els.customerFilterStartDate.value || "";
  let end = els.customerFilterEndDate.value || "";

  if (start && end && start > end) {
    const temp = start;
    start = end;
    end = temp;
  }

  customerFilterState = {
    ...customerFilterState,
    start,
    end
  };
  saveCustomerFilterState();
  updateCustomerDateRangeTrigger();
}

function getCustomerDatePresetRange(preset) {
  const now = new Date();

  if (preset === "today") {
    const date = toDateKey(now);
    return { start: date, end: date };
  }

  if (preset === "7d") {
    const endDate = new Date(now);
    const startDate = new Date(now);
    startDate.setDate(startDate.getDate() - 6);
    return { start: toDateKey(startDate), end: toDateKey(endDate) };
  }

  if (preset === "thisMonth") {
    const startDate = new Date(now.getFullYear(), now.getMonth(), 1);
    const endDate = new Date(now);
    return { start: toDateKey(startDate), end: toDateKey(endDate) };
  }

  if (preset === "lastMonth") {
    const firstDayPrevMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const lastDayPrevMonth = new Date(now.getFullYear(), now.getMonth(), 0);
    return { start: toDateKey(firstDayPrevMonth), end: toDateKey(lastDayPrevMonth) };
  }

  return null;
}

function applyCustomerDatePreset(preset) {
  const range = getCustomerDatePresetRange(preset);
  if (!range) return;
  customerFilterState = {
    ...customerFilterState,
    start: range.start,
    end: range.end
  };
  saveCustomerFilterState();
  els.customerFilterStartDate.value = range.start;
  els.customerFilterEndDate.value = range.end;
  updateCustomerDateRangeTrigger();
}

function setCustomerFilterSummary(filteredCount, totalCount) {
  if (!els.customerFilterSummary) return;
  const datePart = customerFilterState.start && customerFilterState.end
    ? `${customerFilterState.start} đến ${customerFilterState.end}`
    : customerFilterState.start
      ? `từ ${customerFilterState.start}`
      : customerFilterState.end
        ? `đến ${customerFilterState.end}`
        : "mọi ngày";

  const ownerPart = customerFilterState.owner === "all" ? "tất cả người phụ trách" : customerFilterState.owner;
  const statusPart = customerFilterState.status === "all" ? "tất cả trạng thái" : customerFilterState.status;
  const sourcePart = customerFilterState.source === "all" ? "tất cả nguồn" : customerFilterState.source;
  const keywordPart = customerFilterState.keyword ? ` | Từ khóa: ${customerFilterState.keyword}` : "";

  els.customerFilterSummary.textContent = `Hiển thị ${filteredCount}/${totalCount} khách hàng | Ngày: ${datePart} | Phụ trách: ${ownerPart} | Trạng thái: ${statusPart} | Nguồn: ${sourcePart}${keywordPart}`;
}

function applyCustomerFiltersFromInputs() {
  applyCustomerDateRangeFromInputs();
  customerFilterState = {
    start: customerFilterState.start,
    end: customerFilterState.end,
    owner: els.customerFilterOwner.value || "all",
    status: els.customerFilterStatus.value || "all",
    source: els.customerFilterSource.value || "all",
    keyword: (els.customerQuickSearch.value || "").trim()
  };
  saveCustomerFilterState();
}

function resetCustomerFilters() {
  customerFilterState = { start: "", end: "", owner: "all", status: "all", source: "all", keyword: "" };
  saveCustomerFilterState();
}

function setCustomerKeywordFilter(value) {
  customerFilterState = {
    ...customerFilterState,
    keyword: value.trim()
  };
  saveCustomerFilterState();
}

function getFilteredCustomers() {
  const keyword = (customerFilterState.keyword || "").toLowerCase();
  return customers.filter((customer) => {
    const customerDate = getCustomerDateKey(customer);
    if (customerFilterState.start && customerDate < customerFilterState.start) return false;
    if (customerFilterState.end && customerDate > customerFilterState.end) return false;
    if (customerFilterState.owner !== "all" && customer.owner !== customerFilterState.owner) return false;
    if (customerFilterState.status !== "all" && customer.status !== customerFilterState.status) return false;
    if (customerFilterState.source !== "all" && customer.source !== customerFilterState.source) return false;
    if (keyword) {
      const haystack = `${customer.name || ""} ${customer.contactPerson || ""} ${customer.phone || ""} ${customer.email || ""}`.toLowerCase();
      if (!haystack.includes(keyword)) return false;
    }
    return true;
  });
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function toCsvValue(value) {
  const text = String(value ?? "").replace(/"/g, '""');
  return `"${text}"`;
}

function exportFilteredCustomersToExcel() {
  const rows = getFilteredCustomers().sort((a, b) => b.updatedAt - a.updatedAt);
  if (rows.length === 0) {
    els.customerStatusMessage.textContent = "Không có dữ liệu để xuất Excel theo bộ lọc hiện tại.";
    return;
  }

  const header = ["Khách hàng", "Người liên hệ", "Số điện thoại", "Email", "Địa chỉ", "Phân loại", "Người phụ trách", "Trạng thái", "Nguồn data", "Nhu cầu", "Ghi chú", "Cập nhật"];
  const records = rows.map((item) => [
    item.name,
    item.contactPerson,
    item.phone,
    item.email,
    item.address,
    item.tier,
    item.owner,
    item.status,
    item.source,
    item.demand,
    item.note,
    new Date(item.updatedAt).toLocaleString("vi-VN", { hour12: false })
  ]);

  const csv = [header, ...records]
    .map((line) => line.map((cell) => toCsvValue(cell)).join(","))
    .join("\n");

  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8;" });
  downloadBlob(blob, `khach-hang-loc-${today}.csv`);
  showToast(`Đã xuất Excel (CSV) ${rows.length} khách hàng.`);
  logActivity("Khách hàng", "Xuất Excel khách hàng", `Số bản ghi: ${rows.length}`);
  renderActivityTable();
}

async function exportFilteredCustomersToPdf() {
  if (!can("canExportPdf")) {
    els.customerStatusMessage.textContent = "Bạn không có quyền xuất PDF khách hàng.";
    return;
  }

  const rows = getFilteredCustomers();
  if (rows.length === 0) {
    els.customerStatusMessage.textContent = "Không có dữ liệu để xuất PDF theo bộ lọc hiện tại.";
    return;
  }

  showToast("Đang tạo PDF khách hàng...", "info");
  const canvas = await html2canvas(els.customerSection, { scale: 2, useCORS: true });
  const image = canvas.toDataURL("image/png");
  const pdf = new jsPDF("p", "mm", "a4");
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const imageHeight = (canvas.height * pageWidth) / canvas.width;

  let heightLeft = imageHeight;
  let position = 0;

  pdf.addImage(image, "PNG", 0, position, pageWidth, imageHeight);
  heightLeft -= pageHeight;

  while (heightLeft > 0) {
    position -= pageHeight;
    pdf.addPage();
    pdf.addImage(image, "PNG", 0, position, pageWidth, imageHeight);
    heightLeft -= pageHeight;
  }

  pdf.save(`khach-hang-loc-${today}.pdf`);
  showToast(`Đã xuất PDF ${rows.length} khách hàng.`);
  logActivity("Khách hàng", "Xuất PDF khách hàng", `Số bản ghi: ${rows.length}`);
  renderActivityTable();
}

// ── Schedule (Lịch khách hàng) ──────────────────────────────────────────────

const SCHEDULE_STATUS_LABELS = {
  pending: "Chờ xác nhận",
  confirmed: "Đã xác nhận",
  completed: "Đã hoàn thành",
  cancelled: "Đã hủy"
};

const SCHEDULE_LABEL_TO_STATUS = {
  "Chờ xác nhận": "pending",
  "Đã xác nhận": "confirmed",
  "Đã hoàn thành": "completed",
  "Đã hủy": "cancelled"
};

function getScheduleStaffNames() {
  return Array.from(new Set(users.map((user) => (user.fullName || "").trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b, "vi"));
}

function renderScheduleStaffControls() {
  const names = getScheduleStaffNames();
  const staffOptions = names.map((name) => `<option value="${name}">${name}</option>`).join("");

  if (els.scheduleFilterStaff) {
    const selected = scheduleFilterState.staff || "all";
    els.scheduleFilterStaff.innerHTML = `<option value="all">Tất cả nhân sự phụ trách</option>${staffOptions}`;
    els.scheduleFilterStaff.value = names.includes(selected) ? selected : "all";
  }

  if (els.scheduleConsultant) {
    const selected = els.scheduleConsultant.value;
    els.scheduleConsultant.innerHTML = `<option value="">Chọn tư vấn viên</option>${staffOptions}`;
    els.scheduleConsultant.value = names.includes(selected) ? selected : "";
  }

  if (els.scheduleNurse) {
    const selected = els.scheduleNurse.value;
    els.scheduleNurse.innerHTML = `<option value="">Chọn điều dưỡng</option>${staffOptions}`;
    els.scheduleNurse.value = names.includes(selected) ? selected : "";
  }

  if (els.scheduleSale) {
    const selected = els.scheduleSale.value;
    els.scheduleSale.innerHTML = `<option value="">Chọn telesale</option>${staffOptions}`;
    els.scheduleSale.value = names.includes(selected) ? selected : "";
  }
}

const SCHEDULE_INLINE_EDIT_META = {
  registrationDate: { type: "date" },
  appointmentTime: { type: "text" },
  customerName: { type: "text" },
  phone: { type: "text" },
  address: { type: "text" },
  status: { type: "select", options: ["pending", "confirmed", "completed", "cancelled"] },
  note: { type: "text" },
  service: { type: "text" },
  consultant: { type: "text" },
  nurse: { type: "text" },
  source: { type: "text" },
  contractAmount: { type: "number" },
  saleStaff: { type: "text" }
};

function openScheduleInlineEditor(cell, record, field) {
  let meta = SCHEDULE_INLINE_EDIT_META[field];
  if (field === "consultant" || field === "nurse" || field === "saleStaff") {
    meta = { type: "select", options: getScheduleStaffNames() };
  }
  if (!meta) return;
  if (cell.dataset.editing === "true") return;
  cell.dataset.editing = "true";

  const currentValue = field === "note"
    ? (record.note || record.motherCondition || "")
    : (record[field] || "");

  let editor;
  if (meta.type === "select") {
    editor = document.createElement("select");
    editor.className = "schedule-inline-input";
    if (field === "status") {
      editor.innerHTML = meta.options.map((value) => `<option value="${value}"${record.status === value ? " selected" : ""}>${SCHEDULE_STATUS_LABELS[value]}</option>`).join("");
    } else {
      editor.innerHTML = [`<option value="">-- Chọn --</option>`]
        .concat(meta.options.map((value) => `<option value="${value}"${currentValue === value ? " selected" : ""}>${value}</option>`))
        .join("");
    }
  } else {
    editor = document.createElement("input");
    editor.className = "schedule-inline-input";
    editor.type = meta.type;
    editor.value = currentValue;
  }

  cell.innerHTML = "";
  cell.appendChild(editor);
  editor.focus();
  if (editor instanceof HTMLInputElement && editor.type !== "date") {
    editor.setSelectionRange(editor.value.length, editor.value.length);
  }

  const commit = () => {
    if (cell.dataset.editing !== "true") return;
    let newValue = editor instanceof HTMLSelectElement ? editor.value : editor.value.trim();
    if (field === "status") {
      record.status = newValue;
    } else if (field === "contractAmount") {
      record.contractAmount = Number(newValue) || 0;
    } else if (field === "note") {
      record.note = newValue;
    } else {
      record[field] = newValue;
    }
    record.updatedAt = Date.now();
    saveJSON(STORAGE.schedule, schedules);
    cell.dataset.editing = "false";
    renderScheduleTable();
    renderCustomerCarePage();
  };

  const cancel = () => {
    if (cell.dataset.editing !== "true") return;
    cell.dataset.editing = "false";
    renderScheduleTable();
  };

  editor.addEventListener("blur", commit);
  editor.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      commit();
    } else if (event.key === "Escape") {
      event.preventDefault();
      cancel();
    }
  });
}

function fmtMoney(n) {
  if (!n || n === 0) return "";
  return Number(n).toLocaleString("vi-VN") + "đ";
}

function formatScheduleDisplayDate(dateValue) {
  if (!dateValue) return "";
  const [year, month, day] = String(dateValue).split("-");
  if (!year || !month || !day) return String(dateValue);
  return `${day}/${month}`;
}

function getFilteredSchedules() {
  const { month, status, staff, source, keyword } = scheduleFilterState;
  return schedules.filter((s) => {
    if (month && !s.registrationDate.startsWith(month)) return false;
    if (status && s.status !== status) return false;
    if (staff && staff !== "all") {
      if (s.consultant !== staff && s.nurse !== staff && s.saleStaff !== staff) return false;
    }
    if (source && !(s.source || "").toLowerCase().includes(source.toLowerCase())) return false;
    if (keyword) {
      const q = keyword.toLowerCase();
      if (!(s.customerName || "").toLowerCase().includes(q) &&
          !(s.phone || "").toLowerCase().includes(q) &&
          !(s.address || "").toLowerCase().includes(q)) return false;
    }
    return true;
  }).sort((a, b) => a.registrationDate.localeCompare(b.registrationDate) || (a.appointmentTime || "").localeCompare(b.appointmentTime || ""));
}

function renderScheduleTable() {
  if (!els.scheduleBody) return;
  if (els.scheduleFilterMonth && !els.scheduleFilterMonth.value) {
    els.scheduleFilterMonth.value = scheduleFilterState.month;
  }
  const rows = getFilteredSchedules();
  if (els.scheduleFilterSummary) {
    const total = rows.length;
    const contractSum = rows.reduce((s, r) => s + (Number(r.contractAmount) || 0), 0);
    const contractCount = rows.filter((r) => Number(r.contractAmount) > 0).length;
    els.scheduleFilterSummary.textContent = `${total} lịch · ${contractCount} có HĐ · Tổng HĐ: ${Number(contractSum).toLocaleString("vi-VN")}đ`;
  }
  let lastDate = "";
  let counter = 0;
  els.scheduleBody.innerHTML = rows.map((s) => {
    counter += 1;
    const rowClass = `schedule-row-${s.status || "pending"}`;
    const daySeparator = lastDate && s.registrationDate !== lastDate
      ? `<tr class="schedule-day-separator"><td colspan="14"></td></tr>`
      : "";
    lastDate = s.registrationDate;
    const noteValue = s.note || s.motherCondition || "";
    return `${daySeparator}<tr class="${rowClass}" data-sch-id="${s.id}">
      <td class="schedule-col-stt" style="text-align:center;">${counter}</td>
      <td class="schedule-editable schedule-col-date" data-field="registrationDate">${formatScheduleDisplayDate(s.registrationDate)}</td>
      <td class="schedule-editable schedule-col-time" data-field="appointmentTime">${s.appointmentTime || ""}</td>
      <td class="schedule-editable schedule-col-name" data-field="customerName" style="font-weight:600;">${s.customerName || ""}</td>
      <td class="schedule-editable schedule-col-phone" data-field="phone">${s.phone || ""}</td>
      <td class="schedule-editable schedule-col-address" data-field="address">${s.address || ""}</td>
      <td class="schedule-editable schedule-col-status" data-field="status">${SCHEDULE_STATUS_LABELS[s.status] || s.status || ""}</td>
      <td class="schedule-editable schedule-col-note" data-field="note">${noteValue}</td>
      <td class="schedule-editable schedule-col-service" data-field="service">${s.service || ""}</td>
      <td class="schedule-editable schedule-col-consultant" data-field="consultant">${s.consultant || ""}</td>
      <td class="schedule-editable schedule-col-nurse" data-field="nurse">${s.nurse || ""}</td>
      <td class="schedule-editable schedule-col-source" data-field="source">${s.source || ""}</td>
      <td class="schedule-editable schedule-col-contract" data-field="contractAmount">${Number(s.contractAmount || 0).toLocaleString("vi-VN")}</td>
      <td class="schedule-editable schedule-col-sale" data-field="saleStaff">${s.saleStaff || ""}</td>
    </tr>`;
  }).join("");
  syncScheduleBottomScrollerWidth();
}

function openScheduleModal(id) {
  renderScheduleStaffControls();
  editingScheduleId = id || null;
  if (id) {
    const s = schedules.find((x) => x.id === id);
    if (!s) return;
    els.scheduleModalTitle.textContent = `Chỉnh sửa: ${s.customerName}`;
    els.scheduleRegDate.value = s.registrationDate || "";
    els.scheduleTime.value = s.appointmentTime || "";
    els.scheduleStatus.value = s.status || "pending";
    els.scheduleName.value = s.customerName || "";
    els.schedulePhone.value = s.phone || "";
    els.scheduleMotherAge.value = s.motherAge || "";
    els.scheduleAddress.value = s.address || "";
    els.scheduleBirthHistory.value = s.birthHistory || "";
    els.scheduleBabyBirthday.value = s.babyBirthday || "";
    els.scheduleStage.value = s.stage || "";
    els.scheduleService.value = s.service || "";
    els.scheduleMotherCondition.value = s.motherCondition || "";
    els.scheduleBabyCondition.value = s.babyCondition || "";
    els.scheduleConsultant.value = s.consultant || "";
    els.scheduleNurse.value = s.nurse || "";
    els.scheduleSale.value = s.saleStaff || "";
    els.scheduleExpPrice.value = s.experiencePrice || "";
    els.scheduleSessionDuration.value = s.sessionDuration || "";
    els.scheduleSource.value = s.source || "";
    els.scheduleContractAmount.value = s.contractAmount || "";
    els.scheduleNote.value = s.note || "";
  } else {
    els.scheduleModalTitle.textContent = "Thêm lịch trải nghiệm";
    els.scheduleRegDate.value = today;
    els.scheduleTime.value = "";
    els.scheduleStatus.value = "pending";
    els.scheduleName.value = "";
    els.schedulePhone.value = "";
    els.scheduleMotherAge.value = "";
    els.scheduleAddress.value = "";
    els.scheduleBirthHistory.value = "";
    els.scheduleBabyBirthday.value = "";
    els.scheduleStage.value = "";
    els.scheduleService.value = "";
    els.scheduleMotherCondition.value = "";
    els.scheduleBabyCondition.value = "";
    els.scheduleConsultant.value = "";
    els.scheduleNurse.value = "";
    els.scheduleSale.value = "";
    els.scheduleExpPrice.value = "";
    els.scheduleSessionDuration.value = "";
    els.scheduleSource.value = "";
    els.scheduleContractAmount.value = "";
    els.scheduleNote.value = "";
  }
  els.scheduleModal.classList.remove("hidden");
}

function closeScheduleModal() {
  els.scheduleModal.classList.add("hidden");
  editingScheduleId = null;
}

function exportScheduleExcel() {
  const rows = getFilteredSchedules();
  const header = ["STT","Ngày trải nghiệm","Giờ trải nghiệm","Họ tên","Số điện thoại","Địa chỉ","Tình trạng","Ghi chú","Dịch vụ đăng ký","Tư vấn","Điều dưỡng","Nguồn data","Giá trị hợp đồng","Telesale"];
  const records = rows.map((s, idx) => [
    idx + 1,
    s.registrationDate,
    s.appointmentTime,
    s.customerName,
    s.phone,
    s.address,
    SCHEDULE_STATUS_LABELS[s.status] || s.status,
    s.note || s.motherCondition || "",
    s.service,
    s.consultant,
    s.nurse,
    s.source,
    Number(s.contractAmount || 0),
    s.saleStaff
  ]);
  const csv = [header, ...records].map((line) => line.map((cell) => toCsvValue(String(cell||""))).join(",")).join("\n");
  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8;" });
  downloadBlob(blob, `lich-khach-hang-${scheduleFilterState.month || today.slice(0,7)}.csv`);
  showToast(`Đã xuất Excel ${rows.length} lịch.`);
  logActivity("Lịch KH", "Xuất Excel", `Số lịch: ${rows.length}`);
}

async function exportSchedulePdf() {
  const rows = getFilteredSchedules();
  if (rows.length === 0) { showToast("Không có dữ liệu để xuất PDF.", "warning"); return; }
  showToast("Đang tạo PDF lịch...", "info");
  const canvas = await html2canvas(els.scheduleSection, { scale: 1.5, useCORS: true });
  const image = canvas.toDataURL("image/png");
  const pdf = new jsPDF("l", "mm", "a4");
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const imageHeight = (canvas.height * pageWidth) / canvas.width;
  let heightLeft = imageHeight;
  let position = 0;
  pdf.addImage(image, "PNG", 0, position, pageWidth, imageHeight);
  heightLeft -= pageHeight;
  while (heightLeft > 0) {
    position -= pageHeight;
    pdf.addPage();
    pdf.addImage(image, "PNG", 0, position, pageWidth, imageHeight);
    heightLeft -= pageHeight;
  }
  pdf.save(`lich-khach-hang-${scheduleFilterState.month || today.slice(0,7)}.pdf`);
  showToast("Đã xuất PDF lịch thành công.");
  logActivity("Lịch KH", "Xuất PDF", `Số lịch: ${rows.length}`);
}

function getClosedSchedulesForCare() {
  return schedules.filter((item) => Number(item.contractAmount || 0) > 0 && item.status !== "cancelled");
}

function getCustomerCareRows() {
  const map = new Map();
  getClosedSchedulesForCare().forEach((item) => {
    const key = `${String(item.phone || "").trim()}|${String(item.customerName || "").trim().toLowerCase()}`;
    const prev = map.get(key);
    const itemDate = item.registrationDate || "";
    const prevDate = prev?.closedDate || "";
    const shouldReplace = !prev || itemDate > prevDate || (itemDate === prevDate && Number(item.updatedAt || 0) > Number(prev.updatedAt || 0));
    if (shouldReplace) {
      map.set(key, {
        key,
        sourceScheduleId: item.id,
        closedDate: item.registrationDate || "",
        customerName: item.customerName || "",
        phone: item.phone || "",
        address: item.address || "",
        service: item.service || "",
        consultant: item.consultant || "",
        nurse: item.nurse || "",
        saleStaff: item.saleStaff || "",
        source: item.source || "",
        contractAmount: Number(item.contractAmount || 0),
        experiencePrice: Number(item.experiencePrice || 0),
        updatedAt: Number(item.updatedAt || 0)
      });
    }
  });
  return Array.from(map.values()).sort((a, b) => (b.closedDate || "").localeCompare(a.closedDate || "") || b.updatedAt - a.updatedAt);
}

function getCustomerCareOptionSets(rows) {
  const staffSet = new Set();
  const sourceSet = new Set();
  const statusSet = new Set(CARE_STATUS_OPTIONS);
  rows.forEach((item) => {
    [item.consultant, item.nurse, item.saleStaff].forEach((name) => {
      if (String(name || "").trim()) staffSet.add(String(name || "").trim());
    });
    if (String(item.source || "").trim()) sourceSet.add(String(item.source || "").trim());
    const status = customerCareProgress[item.key]?.status;
    if (status) statusSet.add(status);
  });
  return {
    staffs: Array.from(staffSet).sort((a, b) => a.localeCompare(b, "vi")),
    sources: Array.from(sourceSet).sort((a, b) => a.localeCompare(b, "vi")),
    statuses: Array.from(statusSet)
  };
}

function inferCareTotalSessions(row) {
  if (row.experiencePrice > 0 && row.contractAmount > 0) {
    return Math.max(1, Math.round(row.contractAmount / row.experiencePrice));
  }
  return 1;
}

function getCareProgressForRow(row) {
  const saved = customerCareProgress[row.key] || {};
  const totalSessions = Math.max(1, Number(saved.totalSessions || inferCareTotalSessions(row) || 1));
  const usedSessions = Math.max(0, Math.min(totalSessions, Number(saved.usedSessions || 0)));
  return {
    totalSessions,
    usedSessions,
    status: saved.status || "Đang chăm sóc",
    nextDate: saved.nextDate || "",
    note: saved.note || ""
  };
}

function renderCustomerCareFilterControls() {
  if (!els.careFilterStaff || !els.careFilterSource || !els.careFilterStatus) return;
  const rows = getCustomerCareRows();
  const { staffs, sources, statuses } = getCustomerCareOptionSets(rows);

  els.careFilterStaff.innerHTML = ["<option value=\"all\">Tất cả nhân sự</option>"]
    .concat(staffs.map((name) => `<option value="${name}">${name}</option>`))
    .join("");
  els.careFilterSource.innerHTML = ["<option value=\"all\">Tất cả nguồn</option>"]
    .concat(sources.map((name) => `<option value="${name}">${name}</option>`))
    .join("");
  els.careFilterStatus.innerHTML = ["<option value=\"all\">Tất cả trạng thái CSKH</option>"]
    .concat(statuses.map((name) => `<option value="${name}">${name}</option>`))
    .join("");

  els.careFilterStartDate.value = customerCareFilterState.start || "";
  els.careFilterEndDate.value = customerCareFilterState.end || "";
  els.careSearch.value = customerCareFilterState.keyword || "";
  els.careFilterStaff.value = staffs.includes(customerCareFilterState.staff) ? customerCareFilterState.staff : "all";
  els.careFilterStatus.value = statuses.includes(customerCareFilterState.status) ? customerCareFilterState.status : "all";
  els.careFilterSource.value = sources.includes(customerCareFilterState.source) ? customerCareFilterState.source : "all";
  els.careFilterProgress.value = customerCareFilterState.progress || "all";
}

function getFilteredCustomerCareRows() {
  const keyword = String(customerCareFilterState.keyword || "").toLowerCase();
  return getCustomerCareRows().filter((row) => {
    const progress = getCareProgressForRow(row);
    if (customerCareFilterState.start && row.closedDate < customerCareFilterState.start) return false;
    if (customerCareFilterState.end && row.closedDate > customerCareFilterState.end) return false;
    if (customerCareFilterState.staff !== "all") {
      const isMatchStaff = [row.consultant, row.nurse, row.saleStaff].some((name) => name === customerCareFilterState.staff);
      if (!isMatchStaff) return false;
    }
    if (customerCareFilterState.status !== "all" && progress.status !== customerCareFilterState.status) return false;
    if (customerCareFilterState.source !== "all" && row.source !== customerCareFilterState.source) return false;
    if (customerCareFilterState.progress === "not_started" && progress.usedSessions !== 0) return false;
    if (customerCareFilterState.progress === "in_progress" && !(progress.usedSessions > 0 && progress.usedSessions < progress.totalSessions)) return false;
    if (customerCareFilterState.progress === "completed" && progress.usedSessions < progress.totalSessions) return false;
    if (keyword) {
      const haystack = `${row.customerName} ${row.phone} ${row.address} ${row.service} ${progress.note}`.toLowerCase();
      if (!haystack.includes(keyword)) return false;
    }
    return true;
  });
}

function renderCustomerCareTable() {
  if (!els.careBody || !els.careFilterSummary) return;
  const allRows = getCustomerCareRows();
  const rows = getFilteredCustomerCareRows();

  if (!rows.length) {
    els.careBody.innerHTML = '<tr><td colspan="16" style="text-align:center;">Không có khách đã chốt phù hợp với bộ lọc.</td></tr>';
  } else {
    els.careBody.innerHTML = rows.map((row) => {
      const progress = getCareProgressForRow(row);
      const remaining = Math.max(0, progress.totalSessions - progress.usedSessions);
      return `
        <tr>
          <td>${row.closedDate || "--"}</td>
          <td><strong>${row.customerName || "--"}</strong><div class="muted" style="font-size:0.75rem;">${row.address || ""}</div></td>
          <td>${row.phone || "--"}</td>
          <td>${row.service || "--"}</td>
          <td>${row.consultant || "--"}</td>
          <td>${row.nurse || "--"}</td>
          <td>${row.saleStaff || "--"}</td>
          <td>${row.source || "--"}</td>
          <td>${Number(row.contractAmount || 0).toLocaleString("vi-VN")} đ</td>
          <td><input class="care-inline-input care-inline-number" type="number" min="1" step="1" value="${progress.totalSessions}" data-care-key="${row.key}" data-care-field="totalSessions" /></td>
          <td><input class="care-inline-input care-inline-number" type="number" min="0" step="1" value="${progress.usedSessions}" data-care-key="${row.key}" data-care-field="usedSessions" /></td>
          <td class="care-remaining-cell">${remaining}</td>
          <td>
            <select class="care-inline-input" data-care-key="${row.key}" data-care-field="status">
              ${CARE_STATUS_OPTIONS.map((status) => `<option value="${status}"${progress.status === status ? " selected" : ""}>${status}</option>`).join("")}
            </select>
          </td>
          <td><input class="care-inline-input" type="date" value="${progress.nextDate || ""}" data-care-key="${row.key}" data-care-field="nextDate" /></td>
          <td><input class="care-inline-input" type="text" value="${progress.note || ""}" data-care-key="${row.key}" data-care-field="note" placeholder="Ghi chú" /></td>
          <td><button class="btn secondary care-save-btn" type="button" data-care-key="${row.key}">Lưu</button></td>
        </tr>
      `;
    }).join("");
  }

  const datePart = customerCareFilterState.start || customerCareFilterState.end
    ? `${customerCareFilterState.start || "--"} → ${customerCareFilterState.end || "--"}`
    : "mọi ngày";
  const keywordPart = customerCareFilterState.keyword ? ` | Từ khóa: ${customerCareFilterState.keyword}` : "";
  els.careFilterSummary.textContent = `Hiển thị ${rows.length}/${allRows.length} khách đã chốt | Ngày chốt: ${datePart}${keywordPart}`;
}

function renderCustomerCarePage() {
  renderCustomerCareFilterControls();
  renderCustomerCareTable();
}

function exportFilteredCustomerCareToExcel() {
  const rows = getFilteredCustomerCareRows();
  if (!rows.length) {
    showToast("Không có dữ liệu CSKH để xuất Excel theo bộ lọc hiện tại.", "warning");
    return;
  }

  const header = [
    "Ngày chốt",
    "Khách hàng",
    "SĐT",
    "Địa chỉ",
    "Dịch vụ",
    "Tư vấn",
    "Điều dưỡng",
    "Sale",
    "Nguồn",
    "Giá trị HĐ",
    "Liệu trình đăng ký",
    "Đã sử dụng",
    "Còn lại",
    "Trạng thái CSKH",
    "Hẹn chăm tiếp",
    "Ghi chú CSKH"
  ];

  const records = rows.map((row) => {
    const progress = getCareProgressForRow(row);
    const remaining = Math.max(0, progress.totalSessions - progress.usedSessions);
    return [
      row.closedDate || "",
      row.customerName || "",
      row.phone || "",
      row.address || "",
      row.service || "",
      row.consultant || "",
      row.nurse || "",
      row.saleStaff || "",
      row.source || "",
      Number(row.contractAmount || 0),
      progress.totalSessions,
      progress.usedSessions,
      remaining,
      progress.status,
      progress.nextDate || "",
      progress.note || ""
    ];
  });

  const csv = [header, ...records]
    .map((line) => line.map((cell) => toCsvValue(cell)).join(","))
    .join("\n");
  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8;" });
  downloadBlob(blob, `cham-soc-khach-hang-${today}.csv`);
  showToast(`Đã xuất Excel CSKH ${rows.length} khách.`);
  logActivity("CSKH", "Xuất Excel", `Số khách: ${rows.length}`);
  renderActivityTable();
}

async function exportFilteredCustomerCareToPdf() {
  if (!can("canExportPdf")) {
    showToast("Bạn không có quyền xuất PDF CSKH.", "warning");
    return;
  }
  const rows = getFilteredCustomerCareRows();
  if (!rows.length) {
    showToast("Không có dữ liệu CSKH để xuất PDF theo bộ lọc hiện tại.", "warning");
    return;
  }

  showToast("Đang tạo PDF CSKH...", "info");
  const canvas = await html2canvas(els.customerCareSection, { scale: 1.8, useCORS: true });
  const image = canvas.toDataURL("image/png");
  const pdf = new jsPDF("l", "mm", "a4");
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const imageHeight = (canvas.height * pageWidth) / canvas.width;

  let heightLeft = imageHeight;
  let position = 0;
  pdf.addImage(image, "PNG", 0, position, pageWidth, imageHeight);
  heightLeft -= pageHeight;

  while (heightLeft > 0) {
    position -= pageHeight;
    pdf.addPage();
    pdf.addImage(image, "PNG", 0, position, pageWidth, imageHeight);
    heightLeft -= pageHeight;
  }

  pdf.save(`cham-soc-khach-hang-${today}.pdf`);
  showToast(`Đã xuất PDF CSKH ${rows.length} khách.`);
  logActivity("CSKH", "Xuất PDF", `Số khách: ${rows.length}`);
  renderActivityTable();
}

function saveCustomerCareProgressFromRow(careKey, tableRow) {
  const totalInput = tableRow.querySelector('[data-care-field="totalSessions"]');
  const usedInput = tableRow.querySelector('[data-care-field="usedSessions"]');
  const statusInput = tableRow.querySelector('[data-care-field="status"]');
  const nextDateInput = tableRow.querySelector('[data-care-field="nextDate"]');
  const noteInput = tableRow.querySelector('[data-care-field="note"]');
  if (!totalInput || !usedInput || !statusInput || !nextDateInput || !noteInput) return;

  const totalSessions = Math.max(1, Number(totalInput.value || 1));
  const usedSessions = Math.max(0, Math.min(totalSessions, Number(usedInput.value || 0)));

  customerCareProgress[careKey] = {
    ...(customerCareProgress[careKey] || {}),
    totalSessions,
    usedSessions,
    status: String(statusInput.value || "Đang chăm sóc"),
    nextDate: String(nextDateInput.value || ""),
    note: String(noteInput.value || "").trim(),
    updatedAt: Date.now()
  };
  saveJSON(STORAGE.customerCareProgress, customerCareProgress);
  showToast("Đã lưu tiến độ chăm sóc khách hàng.");
  logActivity("CSKH", "Cập nhật tiến độ", `Khóa KH: ${careKey}`);
  renderCustomerCareTable();
}

// ── End Schedule ─────────────────────────────────────────────────────────────

function getLatestRowsByDepartment(list) {
  const map = new Map();
  list.forEach((row) => {
    const prev = map.get(row.department);
    if (!prev || prev.updatedAt < row.updatedAt) map.set(row.department, row);
  });
  return Array.from(map.values());
}

function updateKPI(filteredReports) {
  const latest = getLatestRowsByDepartment(filteredReports);
  const avgCompletion = latest.length ? latest.reduce((s, r) => s + r.completion, 0) / latest.length : 0;
  const avgQuality = latest.length ? latest.reduce((s, r) => s + r.quality, 0) / latest.length : 0;
  const totalIssues = latest.reduce((s, r) => s + r.issues, 0);

  els.kpiReports.textContent = `${latest.length}/${DEPARTMENTS.length}`;
  els.kpiCompletion.textContent = `${avgCompletion.toFixed(1)}%`;
  els.kpiQuality.textContent = avgQuality.toFixed(1);
  els.kpiIssues.textContent = totalIssues;
}

function renderMetricsFilterControls() {
  if (!els.metricsStartDate || !els.metricsEndDate || !els.metricsDepartmentFilter) return;
  els.metricsStartDate.value = metricsFilterState.start || "";
  els.metricsEndDate.value = metricsFilterState.end || "";
  els.metricsDepartmentFilter.value = metricsFilterState.department || "all";
}

function getFilteredSchedulesForMetrics() {
  const start = metricsFilterState.start;
  const end = metricsFilterState.end;
  const dep = metricsFilterState.department;

  const matchesDepartment = (item) => {
    if (dep === "all") return true;
    if (dep === "marketing") {
      const source = String(item.source || "").toLowerCase();
      return ["facebook", "tiktok", "google", "website", "zalo", "instagram", "ads"].some((key) => source.includes(key));
    }
    if (dep === "telesale") return Boolean(String(item.saleStaff || "").trim());
    if (dep === "consultant") return Boolean(String(item.consultant || "").trim());
    if (dep === "nurse") return Boolean(String(item.nurse || "").trim());
    return true;
  };

  return schedules.filter((item) => {
    const d = item.registrationDate || "";
    if (start && d < start) return false;
    if (end && d > end) return false;
    if (!matchesDepartment(item)) return false;
    return true;
  });
}

function buildMetricsFromRows(rows) {
  const marketingSources = ["facebook", "tiktok", "google", "website", "zalo", "instagram", "ads", "marketing"];
  const marketingRows = rows.filter((item) => {
    const source = String(item.source || "").toLowerCase();
    return marketingSources.some((key) => source.includes(key));
  });
  const marketingPhones = new Set(marketingRows.map((item) => String(item.phone || "").trim()).filter(Boolean));
  const marketingLeadRate = rows.length ? (marketingRows.length / rows.length) * 100 : 0;

  const telesaleAssigned = rows.filter((item) => String(item.saleStaff || "").trim());
  const telesaleBooked = telesaleAssigned.filter((item) => item.status === "confirmed" || item.status === "completed");
  const telesaleCanceled = telesaleAssigned.filter((item) => item.status === "cancelled");
  const telesaleRate = telesaleAssigned.length ? (telesaleBooked.length / telesaleAssigned.length) * 100 : 0;
  const telesaleCancelRate = telesaleAssigned.length ? (telesaleCanceled.length / telesaleAssigned.length) * 100 : 0;

  const consultantAssigned = rows.filter((item) => String(item.consultant || "").trim());
  const consultantSigned = consultantAssigned.filter((item) => Number(item.contractAmount || 0) > 0);
  const consultantRevenue = consultantSigned.reduce((sum, item) => sum + Number(item.contractAmount || 0), 0);
  const consultantSignRate = consultantAssigned.length ? (consultantSigned.length / consultantAssigned.length) * 100 : 0;
  const consultantAvgContract = consultantSigned.length ? consultantRevenue / consultantSigned.length : 0;

  const nurseAssigned = rows.filter((item) => String(item.nurse || "").trim());
  const nurseCaring = rows.filter((item) => String(item.nurse || "").trim() && (item.status === "pending" || item.status === "confirmed"));
  const nurseCompleted = nurseAssigned.filter((item) => item.status === "completed");
  const nurseServiceScore = nurseAssigned.length ? Math.min(5, (nurseCompleted.length / nurseAssigned.length) * 5) : 0;
  const nurseCompleteRate = nurseAssigned.length ? (nurseCompleted.length / nurseAssigned.length) * 100 : 0;

  return {
    marketingInteractions: rows.length,
    marketingPhones: marketingPhones.size,
    marketingLeadRate,
    telesaleAppointments: telesaleAssigned.length,
    telesaleBookingRate: telesaleRate,
    telesaleCancelRate,
    consultantRevenue,
    consultantSignRate,
    consultantAvgContract,
    nurseCaring: nurseCaring.length,
    nurseServiceScore,
    nurseCompleteRate
  };
}

function drawMetricsDeptKpiChart(metrics) {
  if (!els.metricsDeptKpiChart) return;
  const labels = ["Marketing", "Telesale", "Tư vấn", "Điều dưỡng"];
  const values = [
    metrics.marketingLeadRate,
    metrics.telesaleBookingRate,
    metrics.consultantSignRate,
    metrics.nurseCompleteRate
  ];
  if (metricsDeptKpiChart) metricsDeptKpiChart.destroy();
  metricsDeptKpiChart = new Chart(els.metricsDeptKpiChart, {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label: "% hiệu quả",
        data: values,
        backgroundColor: ["#2563eb", "#0ea5e9", "#16a34a", "#f59e0b"],
        borderRadius: 8
      }]
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false } },
      scales: { y: { beginAtZero: true, max: 100 } }
    }
  });
}

function parseCsvLine(line) {
  const result = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === "," && !inQuotes) {
      result.push(current);
      current = "";
    } else {
      current += char;
    }
  }
  result.push(current);
  return result.map((value) => value.trim());
}

function parseCsvText(text) {
  const lines = text.split(/\r?\n/).filter((line) => line.trim());
  if (lines.length < 2) return [];
  const headers = parseCsvLine(lines[0]).map((h) => h.toLowerCase());
  return lines.slice(1).map((line) => {
    const values = parseCsvLine(line);
    const row = {};
    headers.forEach((header, idx) => {
      row[header] = values[idx] || "";
    });
    return row;
  });
}

function firstValue(obj, keys) {
  for (const key of keys) {
    if (obj[key] !== undefined && obj[key] !== null && String(obj[key]).trim() !== "") return obj[key];
  }
  return "";
}

// ─── Telegram Integration ─────────────────────────────────────────────────────
function saveTelegramSourceConfig() {
  saveJSON(STORAGE.telegramSource, telegramSourceConfig);
}

const TELEGRAM_BRIDGE_DEFAULT_API = "http://localhost:8787";
function getTelegramBridgeApiBase() {
  const configured = String(telegramSourceConfig.bridgeApiUrl || "").trim();
  return (configured || TELEGRAM_BRIDGE_DEFAULT_API).replace(/\/+$/, "");
}

async function callTelegramBridge(path, options = {}) {
  const base = getTelegramBridgeApiBase();
  if (!base) throw new Error("Chưa cấu hình Telegram Bridge API URL.");
  const response = await fetch(`${base}${path}`, {
    method: options.method || "GET",
    headers: { "Content-Type": "application/json", ...(options.headers || {}) },
    body: options.body
  });
  const text = await response.text();
  const data = text ? JSON.parse(text) : {};
  if (!response.ok || data.ok === false) {
    throw new Error(data.error || data.message || `HTTP ${response.status}`);
  }
  return data;
}

function parseTelegramNurseMessage(text) {
  if (!text) return null;
  const lower = text.toLowerCase();
  if (!lower.includes("#")) return null;
  const lines = text.split(/[\n\r]+/);
  const obj = {};
  const tags = Array.from(new Set((text.match(/#[^\s#]+/g) || []).map((tag) => tag.replace(/^#/, "").trim().toLowerCase())));
  lines.forEach(line => {
    const sep = line.indexOf(":");
    if (sep === -1) return;
    const key = line.slice(0, sep).trim().toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, "");
    const val = line.slice(sep + 1).trim();
    if (key && val) obj[key] = val;
  });

  const hasTag = (candidates) => candidates.some((tag) => tags.includes(tag));
  const route = hasTag(["mkt", "marketing", "mk"])
    ? "marketing"
    : hasTag(["tuvan", "tv", "consultant"])
      ? "consultant"
      : hasTag(["telesale", "sale", "ts"])
        ? "telesale"
        : hasTag(["dieuduong", "dd", "baocao", "nurse"])
          ? "nurse"
          : "";
  if (!route) return null;

  const base = {
    registrationDate: obj["ngay"] || obj["date"] || "",
    appointmentTime: obj["gio"] || obj["time"] || "",
    customerName: obj["khach"] || obj["khachhang"] || obj["customer"] || "",
    phone: obj["sdt"] || obj["sodienthoai"] || obj["phone"] || "",
    service: obj["dichvu"] || obj["service"] || "",
    shiftMinutes: obj["thoiluong"] || obj["phut"] || obj["minutes"] || "",
    distanceKm: obj["khoangcach"] || obj["km"] || obj["distance"] || "",
    contractAmount: obj["hopdong"] || obj["contract"] || obj["doanhso"] || "",
    status: obj["trangthai"] || obj["status"] || "completed",
    note: obj["ghichu"] || obj["note"] || "",
    telegramTags: tags,
    telegramRoute: route,
    source: "Telegram Webhook"
  };

  if (route === "nurse") {
    return {
      ...base,
      nurse: obj["ten"] || obj["dieuduong"] || obj["nurse"] || "",
      source: "Telegram Webhook #dieuduong"
    };
  }

  if (route === "marketing") {
    const marketingName = obj["ten"] || obj["marketing"] || obj["mkt"] || obj["marketer"] || "";
    return {
      ...base,
      marketingName,
      marketingStaff: marketingName,
      marketingBudget: obj["ngansach"] || obj["budget"] || obj["ads"] || 0,
      marketingMessCount: obj["mess"] || obj["luongmess"] || obj["interactions"] || 1,
      source: "Telegram Marketing #mkt"
    };
  }

  if (route === "consultant") {
    return {
      ...base,
      consultant: obj["ten"] || obj["tuvan"] || obj["consultant"] || obj["tv"] || "",
      source: "Telegram Tu Van #tuvan"
    };
  }

  return {
    ...base,
    saleStaff: obj["ten"] || obj["telesale"] || obj["sale"] || obj["ts"] || "",
    source: "Telegram Telesale #telesale"
  };
}

function parseTelegramAllowedChatIds(chatIdValue) {
  return String(chatIdValue || "")
    .split(/[\s,;|]+/)
    .map((id) => id.trim())
    .filter(Boolean);
}

async function fetchAndParseTelegramMessagesDirect() {
  const { token, chatId, lastUpdateId } = telegramSourceConfig;
  if (!token || !chatId) throw new Error("Chưa nhập Bot Token hoặc Chat ID.");
  const allowedChatIds = parseTelegramAllowedChatIds(chatId);
  const offset = lastUpdateId ? lastUpdateId + 1 : undefined;
  const url = "https://api.telegram.org/bot" + token + "/getUpdates?limit=100" +
    (offset ? "&offset=" + offset : "");
  const res = await fetch(url);
  if (!res.ok) throw new Error("Telegram API lỗi HTTP: " + res.status);
  const data = await res.json();
  if (!data.ok) throw new Error(data.description || "Telegram trả về lỗi.");

  const updates = data.result || [];
  let maxUpdateId = telegramSourceConfig.lastUpdateId;
  const rawRows = [];
  updates.forEach(u => {
    if (!u.message) return;
    const incomingChatId = String(u.message.chat && u.message.chat.id);
    if (!allowedChatIds.includes(incomingChatId)) return;
    if (u.update_id > maxUpdateId) maxUpdateId = u.update_id;
    const parsed = parseTelegramNurseMessage(u.message.text || "");
    if (parsed) rawRows.push({
      ...parsed,
      telegramUpdateId: String(u.update_id || ""),
      telegramChatId: incomingChatId,
      telegramChatTitle: String(u.message.chat?.title || u.message.chat?.username || "").trim()
    });
  });

  const imported = mergeImportedSchedules(rawRows);
  telegramSourceConfig.lastUpdateId = maxUpdateId;
  telegramSourceConfig.lastSyncedAt = Date.now();
  saveTelegramSourceConfig();
  return { totalUpdates: updates.length, importedRows: imported, parsedMessages: rawRows.length };
}

async function configureTelegramRealtimeWebhook() {
  const token = String(telegramSourceConfig.token || "").trim();
  const chatId = String(telegramSourceConfig.chatId || "").trim();
  const webhookBaseUrl = String(telegramSourceConfig.webhookBaseUrl || "").trim();
  if (!token || !chatId) throw new Error("Vui lòng nhập Bot Token và Chat ID.");
  return callTelegramBridge("/api/telegram/config", {
    method: "POST",
    body: JSON.stringify({ token, chatId, webhookBaseUrl })
  });
}

async function pullTelegramWebhookReports(options = {}) {
  const { fullSync = false } = options;
  const since = fullSync ? 0 : Number(telegramSourceConfig.lastSyncedAt || 0);
  const data = await callTelegramBridge(`/api/telegram/pending?since=${since}${fullSync ? "&all=1" : ""}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  const importedRows = fullSync ? replaceTelegramSchedules(rows) : mergeImportedSchedules(rows);
  if (data.lastReceivedAt) {
    telegramSourceConfig.lastSyncedAt = Math.max(Number(telegramSourceConfig.lastSyncedAt || 0), Number(data.lastReceivedAt || 0));
    saveTelegramSourceConfig();
  }
  return {
    importedRows,
    fetchedRows: rows.length,
    pendingCount: Number(data.pendingCount || 0),
    configured: Boolean(data.configured)
  };
}

async function runTelegramRealtimeSync(silent = false, options = {}) {
  try {
    const result = await pullTelegramWebhookReports(options);
    if (!silent) {
      const msg = result.importedRows > 0
        ? `Đã nhập ${result.importedRows} ca từ webhook Telegram.`
        : "Không có báo cáo Telegram mới.";
      showToast(msg, result.importedRows > 0 ? "success" : "info");
    }
    if (result.importedRows > 0) renderAll();
    return result;
  } catch (err) {
    if (silent) return { importedRows: 0, fetchedRows: 0, pendingCount: 0, configured: false, error: err.message };
    throw err;
  }
}
// ─── End Telegram Integration ──────────────────────────────────────────────────

function normalizeImportedScheduleRow(raw) {
  const sourceObj = {};
  Object.keys(raw || {}).forEach((key) => {
    sourceObj[String(key).toLowerCase().trim()] = raw[key];
  });

  const dateRaw = firstValue(sourceObj, ["registrationdate", "date", "ngay", "ngày trải nghiệm", "ngay trai nghiem"]);
  const normalizedDate = String(dateRaw || "").includes("-")
    ? String(dateRaw).slice(0, 10)
    : toDateKey(new Date(dateRaw || Date.now()));

  const statusRaw = String(firstValue(sourceObj, ["status", "tình trạng", "tinh trang"]) || "").toLowerCase();
  const status = statusRaw.includes("hủy") || statusRaw.includes("huy") || statusRaw === "cancelled"
    ? "cancelled"
    : statusRaw.includes("hoàn") || statusRaw.includes("hoan") || statusRaw === "completed"
      ? "completed"
      : statusRaw.includes("xác") || statusRaw.includes("xac") || statusRaw === "confirmed"
        ? "confirmed"
        : "pending";

  const customerName = String(firstValue(sourceObj, ["customername", "họ tên", "ho ten", "name"]) || "").trim();
  if (!customerName) return null;

  const telegramUpdateId = String(firstValue(sourceObj, ["telegramupdateid"]) || "").trim();

  return {
    id: telegramUpdateId ? `tg-${telegramUpdateId}` : `sc-${Date.now()}-${Math.floor(Math.random() * 100000)}`,
    registrationDate: normalizedDate,
    appointmentTime: String(firstValue(sourceObj, ["appointmenttime", "time", "gio", "giờ trải nghiệm", "gio trai nghiem"]) || "").trim(),
    customerName,
    phone: String(firstValue(sourceObj, ["phone", "số điện thoại", "so dien thoai"]) || "").trim(),
    address: String(firstValue(sourceObj, ["address", "địa chỉ", "dia chi"]) || "").trim(),
    motherAge: "",
    birthHistory: "",
    babyBirthday: "",
    priority: "",
    service: String(firstValue(sourceObj, ["service", "dịch vụ", "dich vu"]) || "").trim(),
    stage: "",
    motherCondition: "",
    babyCondition: "",
    consultant: String(firstValue(sourceObj, ["consultant", "tư vấn", "tu van"]) || "").trim(),
    nurse: String(firstValue(sourceObj, ["nurse", "điều dưỡng", "dieu duong"]) || "").trim(),
    saleStaff: String(firstValue(sourceObj, ["telesale", "sale", "telesale phụ trách", "telesale phu trach"]) || "").trim(),
    experiencePrice: Number(firstValue(sourceObj, ["experienceprice", "gia tn"]) || 0),
    sessionDuration: String(firstValue(sourceObj, ["sessionduration", "thời gian", "thoi gian"]) || "").trim(),
    source: String(firstValue(sourceObj, ["source", "nguồn data", "nguon data"]) || "Nhập file").trim(),
    contractAmount: Number(firstValue(sourceObj, ["contractamount", "giá trị hợp đồng", "gia tri hop dong"]) || 0),
    shiftMinutes: Number(firstValue(sourceObj, ["shiftminutes", "minutes", "phut", "số phút", "so phut", "thời lượng ca", "thoi luong ca"]) || 0),
    distanceKm: parseFlexibleNumber(firstValue(sourceObj, ["distancekm", "distance", "khoang cach", "khoảng cách", "quang duong", "quãng đường", "km"])),
    renewContractAmount: Number(firstValue(sourceObj, ["renewcontractamount", "tai ky", "tái ký", "gia tri tai ky", "giá trị tái ký"]) || 0),
    productSalesAmount: Number(firstValue(sourceObj, ["productsalesamount", "doanh so san pham", "doanh số sản phẩm", "ban san pham", "bán sản phẩm"]) || 0),
    hotBonus: Number(firstValue(sourceObj, ["hotbonus", "thuong nong", "thưởng nóng"]) || 0),
    monthlyCompetitionBonus: Number(firstValue(sourceObj, ["monthlycompetitionbonus", "thuong thi dua", "thưởng thi đua", "thuong thi dua thang", "thưởng thi đua tháng"]) || 0),
    socialInsuranceDeduction: Number(firstValue(sourceObj, ["socialinsurancededuction", "bhxh", "khau tru bhxh", "khấu trừ bhxh"]) || 0),
    trainingDeduction: Number(firstValue(sourceObj, ["trainingdeduction", "phi dao tao", "phí đào tạo", "khau tru dao tao", "khấu trừ đào tạo"]) || 0),
    unionDeduction: Number(firstValue(sourceObj, ["uniondeduction", "phi cong doan", "phí công đoàn", "khau tru cong doan", "khấu trừ công đoàn"]) || 0),
    violationDeduction: Number(firstValue(sourceObj, ["violationdeduction", "phat vi pham", "phạt vi phạm", "tien phat", "tiền phạt"]) || 0),
    telegramUpdateId,
    status,
    note: String(firstValue(sourceObj, ["note", "ghi chú", "ghi chu"]) || "").trim(),
    updatedAt: Date.now(),
    createdAt: Date.now()
  };
}

function mergeImportedSchedules(imported) {
  const existingIds = new Set((schedules || []).map((item) => String(item.id || "")));
  const list = imported
    .map(normalizeImportedScheduleRow)
    .filter(Boolean)
    .filter((item) => {
      const key = String(item.id || "");
      if (!key || !existingIds.has(key)) {
        if (key) existingIds.add(key);
        return true;
      }
      return false;
    });
  if (!list.length) return 0;
  schedules = [...list, ...schedules];
  saveJSON(STORAGE.schedule, schedules);
  return list.length;
}

function replaceTelegramSchedules(imported) {
  const normalizedTelegramRows = imported.map(normalizeImportedScheduleRow).filter(Boolean);
  const incomingTelegramIds = new Set(normalizedTelegramRows.map((item) => String(item.id || "")).filter(Boolean));
  const existingTelegramRows = (schedules || []).filter((item) => {
    const id = String(item.id || "");
    const source = String(item.source || "").toLowerCase();
    return id.startsWith("tg-") || source.includes("telegram");
  });
  const existingTelegramIds = new Set(existingTelegramRows.map((item) => String(item.id || "")).filter(Boolean));

  const nonTelegramRows = (schedules || []).filter((item) => {
    const id = String(item.id || "");
    const source = String(item.source || "").toLowerCase();
    return !(id.startsWith("tg-") || source.includes("telegram"));
  });
  schedules = [...normalizedTelegramRows, ...nonTelegramRows];
  saveJSON(STORAGE.schedule, schedules);

  let changedCount = 0;
  incomingTelegramIds.forEach((id) => {
    if (!existingTelegramIds.has(id)) changedCount += 1;
  });
  existingTelegramIds.forEach((id) => {
    if (!incomingTelegramIds.has(id)) changedCount += 1;
  });
  return changedCount;
}

function normalizeSheetUrl(url) {
  if (!url.includes("docs.google.com/spreadsheets")) return url;
  if (url.includes("/gviz/tq") || url.includes("output=csv")) return url;
  const clean = url.split("?")[0];
  return `${clean.replace(/\/edit$/, "")}/gviz/tq?tqx=out:csv`;
}

const OFFICIAL_METRICS_FALLBACK = {
  marketingInteractions: 126,
  marketingPhones: 52,
  marketingLeadRate: 41.2,
  telesaleAppointments: 178,
  telesaleBookingRate: 34.8,
  telesaleCancelRate: 7.2,
  consultantRevenue: 1420000000,
  consultantSignRate: 36.4,
  consultantAvgContract: 22187500,
  nurseCaring: 94,
  nurseServiceScore: 4.7,
  nurseCompleteRate: 63.5
};

function renderDepartmentMetrics() {
  if (!els.mkInteractions) return;

  const rows = getFilteredSchedulesForMetrics();

  if (rows.length === 0) {
    const fallback = OFFICIAL_METRICS_FALLBACK;
    els.mkInteractions.textContent = fallback.marketingInteractions.toLocaleString("vi-VN");
    els.mkPhones.textContent = fallback.marketingPhones.toLocaleString("vi-VN");
    els.mkLeadRate.textContent = `${fallback.marketingLeadRate.toFixed(1)}%`;
    els.tsAppointments.textContent = fallback.telesaleAppointments.toLocaleString("vi-VN");
    els.tsBookingRate.textContent = `${fallback.telesaleBookingRate.toFixed(1)}%`;
    els.tsCancelRate.textContent = `${fallback.telesaleCancelRate.toFixed(1)}%`;
    els.tvRevenue.textContent = `${fallback.consultantRevenue.toLocaleString("vi-VN")} đ`;
    els.tvSignRate.textContent = `${fallback.consultantSignRate.toFixed(1)}%`;
    els.tvAvgContract.textContent = `${fallback.consultantAvgContract.toLocaleString("vi-VN")} đ`;
    els.ddCaring.textContent = fallback.nurseCaring.toLocaleString("vi-VN");
    els.ddServiceScore.textContent = `${fallback.nurseServiceScore.toFixed(1)}/5`;
    els.ddCompleteRate.textContent = `${fallback.nurseCompleteRate.toFixed(1)}%`;
    drawMetricsDeptKpiChart({
      marketingLeadRate: fallback.marketingLeadRate,
      telesaleBookingRate: fallback.telesaleBookingRate,
      consultantSignRate: fallback.consultantSignRate,
      nurseCompleteRate: fallback.nurseCompleteRate
    });
    if (els.metricsFilterSummary) {
      const startLabel = metricsFilterState.start || "--";
      const endLabel = metricsFilterState.end || "--";
      const depMap = { all: "tất cả phòng ban", marketing: "Marketing", telesale: "Telesale", consultant: "Tư vấn", nurse: "Điều dưỡng" };
      const depLabel = depMap[metricsFilterState.department] || "tất cả phòng ban";
      els.metricsFilterSummary.textContent = `Bộ lọc: ${startLabel} → ${endLabel} | ${depLabel} · 0 lịch (đang dùng dữ liệu tạm)`;
    }
    return;
  }
  const metrics = buildMetricsFromRows(rows);
  els.mkInteractions.textContent = metrics.marketingInteractions.toLocaleString("vi-VN");
  els.mkPhones.textContent = metrics.marketingPhones.toLocaleString("vi-VN");
  els.mkLeadRate.textContent = `${metrics.marketingLeadRate.toFixed(1)}%`;
  els.tsAppointments.textContent = metrics.telesaleAppointments.toLocaleString("vi-VN");
  els.tsBookingRate.textContent = `${metrics.telesaleBookingRate.toFixed(1)}%`;
  els.tsCancelRate.textContent = `${metrics.telesaleCancelRate.toFixed(1)}%`;
  els.tvRevenue.textContent = `${metrics.consultantRevenue.toLocaleString("vi-VN")} đ`;
  els.tvSignRate.textContent = `${metrics.consultantSignRate.toFixed(1)}%`;
  els.tvAvgContract.textContent = `${Math.round(metrics.consultantAvgContract).toLocaleString("vi-VN")} đ`;
  els.ddCaring.textContent = metrics.nurseCaring.toLocaleString("vi-VN");
  els.ddServiceScore.textContent = `${metrics.nurseServiceScore.toFixed(1)}/5`;
  els.ddCompleteRate.textContent = `${metrics.nurseCompleteRate.toFixed(1)}%`;
  drawMetricsDeptKpiChart(metrics);

  if (els.metricsFilterSummary) {
    const startLabel = metricsFilterState.start || "--";
    const endLabel = metricsFilterState.end || "--";
    const depMap = { all: "tất cả phòng ban", marketing: "Marketing", telesale: "Telesale", consultant: "Tư vấn", nurse: "Điều dưỡng" };
    const depLabel = depMap[metricsFilterState.department] || "tất cả phòng ban";
    els.metricsFilterSummary.textContent = `Bộ lọc: ${startLabel} → ${endLabel} | ${depLabel} · ${rows.length} lịch`;
  }
}

const REPORT_DEPARTMENT_META = {
  marketing: {
    title: "Báo cáo Marketing",
    desc: "",
    cols: ["Ngân sách", "Lượng mess", "Số điện thoại", "Doanh số"]
  },
  telesale: {
    title: "Báo cáo Telesale",
    desc: "",
    cols: ["Lượng mess", "Lịch trải nghiệm", "Tỉ lệ đặt lịch", "Doanh số"]
  },
  consultant: {
    title: "Báo cáo Tư vấn",
    desc: "",
    cols: ["Số ca nhận", "Số ca ký", "Tỉ lệ ký", "Doanh số"]
  },
  nurse: {
    title: "Báo cáo Điều dưỡng",
    desc: "Danh sách chi tiết theo từng điều dưỡng và số ca đã làm trong tháng.",
    cols: ["Tổng ca tháng", "Ca chăm bé", "Ca chăm mẹ", "Ghi chú dịch vụ"]
  }
};

const NURSE_REPORT_BUCKETS = [
  { key: "baby30", label: "Tắm bé 30P", group: "baby", minutes: 30 },
  { key: "baby45", label: "Tắm bé 45P", group: "baby", minutes: 45 },
  { key: "baby60", label: "Tắm bé 60P", group: "baby", minutes: 60 },
  { key: "mother60", label: "Mẹ 60P", group: "mother", minutes: 60 },
  { key: "mother75", label: "Ca mẹ 75P", group: "mother", minutes: 75 },
  { key: "mother90", label: "Mẹ 90P", group: "mother", minutes: 90 },
  { key: "mother120", label: "Mẹ 120P", group: "mother", minutes: 120 }
];

function toggleNurseReportSort(sortKey) {
  if (nurseReportSortState.key === sortKey) {
    nurseReportSortState = {
      key: sortKey,
      direction: nurseReportSortState.direction === "desc" ? "asc" : "desc"
    };
    return;
  }
  nurseReportSortState = { key: sortKey, direction: "desc" };
}

function getNurseReportSortIndicator(sortKey) {
  if (nurseReportSortState.key !== sortKey) return "";
  return nurseReportSortState.direction === "desc" ? " ↓" : " ↑";
}

function getShiftMinutes(item) {
  const directMinutes = Number(item.shiftMinutes) || 0;
  if (directMinutes > 0) return directMinutes;
  const matched = String(item.sessionDuration || "").match(/(\d{2,3})/);
  return matched ? Number(matched[1]) : 90;
}

function parseFlexibleNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const text = String(value || "").trim();
  if (!text) return 0;
  const normalized = text.replace(/,/g, ".").replace(/[^\d.\-]/g, "");
  const parsed = Number.parseFloat(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function getDistanceKm(item) {
  return Math.max(0, parseFlexibleNumber(item?.distanceKm));
}

function getClosestBucketByMinutes(group, minutes) {
  const candidates = NURSE_REPORT_BUCKETS.filter((bucket) => bucket.group === group);
  return candidates.reduce((closest, bucket) => {
    if (!closest) return bucket;
    const currentDistance = Math.abs(bucket.minutes - minutes);
    const closestDistance = Math.abs(closest.minutes - minutes);
    return currentDistance < closestDistance ? bucket : closest;
  }, null);
}

function normalizeTextForMatching(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function getCanonicalNurseName(name) {
  const rawName = String(name || "").trim();
  if (!rawName) return "";

  const normalizedName = normalizeTextForMatching(rawName);
  const shortNameCandidates = Array.from(new Set(
    (schedules || [])
      .map((item) => String(item.nurse || "").trim())
      .filter((candidate) => candidate && !candidate.includes(" "))
  ));

  const exactCandidate = shortNameCandidates.find((candidate) => normalizeTextForMatching(candidate) === normalizedName);
  if (exactCandidate) return exactCandidate;

  const suffixCandidate = shortNameCandidates.find((candidate) => {
    const normalizedCandidate = normalizeTextForMatching(candidate);
    return normalizedCandidate.length >= 2 && normalizedName.endsWith(normalizedCandidate);
  });
  if (suffixCandidate) return suffixCandidate;

  return rawName;
}

function getNurseServiceBucket(item) {
  const serviceText = String(item.service || "").toLowerCase();
  const stageText = String(item.stage || "").toLowerCase();
  const minutes = getShiftMinutes(item);
  const isBabyCare = serviceText.includes("bé") || serviceText.includes("tam be") || stageText.includes("em bé");
  const group = isBabyCare ? "baby" : "mother";
  return getClosestBucketByMinutes(group, minutes);
}

function getNurseDetailedReportMatrix(start, end) {
  const rows = getRowsByReportDepartment("nurse", start, end)
    .filter((item) => item.status === "completed");
  const bucketByNurse = new Map();

  rows.forEach((item) => {
    const nurseName = getCanonicalNurseName(item.nurse);
    if (!nurseName) return;
    const bucket = getNurseServiceBucket(item);
    if (!bucket) return;

    const nurseRow = bucketByNurse.get(nurseName) || {
      nurseName,
      counts: Object.fromEntries(NURSE_REPORT_BUCKETS.map((column) => [column.key, 0])),
      workingDates: new Set(),
      travelAllowance: 0
    };
    nurseRow.counts[bucket.key] += 1;
    if (item.registrationDate) nurseRow.workingDates.add(item.registrationDate);
    const distanceKm = getDistanceKm(item);
    if (distanceKm > 20) nurseRow.travelAllowance += 20000;
    else if (distanceKm > 15) nurseRow.travelAllowance += 15000;
    else if (distanceKm > 10) nurseRow.travelAllowance += 10000;
    bucketByNurse.set(nurseName, nurseRow);
  });

  const totals = Object.fromEntries(NURSE_REPORT_BUCKETS.map((bucket) => [bucket.key, 0]));
  let totalTravelAllowance = 0;

  const detailRows = Array.from(bucketByNurse.values())
    .map((row) => {
      const totalMinutes = NURSE_REPORT_BUCKETS.reduce((sum, bucket) => {
        const count = Math.max(0, Number(row.counts[bucket.key]) || 0);
        totals[bucket.key] += count;
        return sum + count * bucket.minutes;
      }, 0);
      totalTravelAllowance += row.travelAllowance;
      return {
        nurseName: row.nurseName,
        counts: row.counts,
        workingDays: row.workingDates.size,
        total: NURSE_REPORT_BUCKETS.reduce((sum, bucket) => sum + (Number(row.counts[bucket.key]) || 0), 0),
        totalMinutes,
        standardShiftCount: totalMinutes / 90,
        travelAllowance: row.travelAllowance
      };
    })
    .sort((a, b) => {
      const directionFactor = nurseReportSortState.direction === "asc" ? 1 : -1;
      let compareValue = 0;
      if (nurseReportSortState.key === "nurseName") {
        compareValue = a.nurseName.localeCompare(b.nurseName);
      } else if (nurseReportSortState.key === "workingDays") {
        compareValue = a.workingDays - b.workingDays;
      } else if (nurseReportSortState.key === "totalMinutes") {
        compareValue = a.totalMinutes - b.totalMinutes;
      } else if (nurseReportSortState.key === "standardShiftCount") {
        compareValue = a.standardShiftCount - b.standardShiftCount;
      } else if (nurseReportSortState.key === "travelAllowance") {
        compareValue = a.travelAllowance - b.travelAllowance;
      } else if (nurseReportSortState.key in a.counts) {
        compareValue = (Number(a.counts[nurseReportSortState.key]) || 0) - (Number(b.counts[nurseReportSortState.key]) || 0);
      } else {
        compareValue = a.total - b.total;
      }
      if (compareValue === 0) return a.nurseName.localeCompare(b.nurseName);
      return compareValue * directionFactor;
    });

  const totalWorkedMinutes = detailRows.reduce((sum, row) => sum + row.totalMinutes, 0);
  const totalWorkingDays = detailRows.reduce((sum, row) => sum + row.workingDays, 0);
  const overallTotal = NURSE_REPORT_BUCKETS.reduce((sum, bucket) => sum + totals[bucket.key], 0);
  return {
    rows: detailRows,
    totals,
    overallTotal,
    totalWorkingDays,
    totalWorkedMinutes,
    totalStandardShiftCount: totalWorkedMinutes / 90,
    totalTravelAllowance
  };
}

function showNurseDailyDetailModal(nurseName, start, end) {
  const shifts = schedules.filter((item) => {
    const name = getCanonicalNurseName(item.nurse);
    if (name !== nurseName) return false;
    if (item.status !== "completed") return false;
    const d = item.registrationDate || "";
    if (start && d < start) return false;
    if (end && d > end) return false;
    return true;
  });

  shifts.sort((a, b) => {
    const da = a.registrationDate || "";
    const db = b.registrationDate || "";
    if (da < db) return -1;
    if (da > db) return 1;
    return (a.appointmentTime || "").localeCompare(b.appointmentTime || "");
  });

  const tbody = shifts.length
    ? shifts.map((item, index) => {
        const bucket = getNurseServiceBucket(item);
        const bucketLabel = bucket ? bucket.label : (item.service || "—");
        const minutes = getShiftMinutes(item);
        const dateFormatted = (() => {
          const d = item.registrationDate || "";
          if (!d) return "—";
          const parts = d.split("-");
          return parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : d;
        })();
        return `<tr>
          <td style="text-align:center;">${index + 1}</td>
          <td style="text-align:center;">${dateFormatted}</td>
          <td style="text-align:center;">${item.appointmentTime || "—"}</td>
          <td>${item.customerName || "—"}</td>
          <td>${item.service || "—"}</td>
          <td>${bucketLabel}</td>
          <td style="text-align:center;">${minutes}</td>
        </tr>`;
      }).join("")
    : `<tr><td colspan="7" style="text-align:center;padding:16px 8px;">Không có ca hoàn tất trong khoảng ngày đã chọn.</td></tr>`;

  const totalMinutes = shifts.reduce((sum, item) => sum + getShiftMinutes(item), 0);
  const periodLabel = start && end ? `${start} → ${end}` : (start || end || "Tất cả");

  const modalEl = document.createElement("div");
  modalEl.id = "nurseDailyDetailModal";
  modalEl.style.cssText = "position:fixed;inset:0;z-index:9999;display:flex;align-items:flex-start;justify-content:center;background:rgba(5,18,30,0.5);overflow-y:auto;padding:24px 12px;";
  modalEl.innerHTML = `
    <div style="position:relative;width:min(860px,94vw);background:#fff;border-radius:16px;box-shadow:0 8px 40px rgba(0,0,0,0.18);padding:20px 22px 24px;">
      <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:14px;">
        <div>
          <h3 style="margin:0;font-size:1.08rem;">Công ca chi tiết — ${nurseName}</h3>
          <p style="margin:4px 0 0;font-size:0.82rem;color:#6b7a99;">Khoảng thời gian: ${periodLabel} &nbsp;|&nbsp; ${shifts.length} ca hoàn tất &nbsp;|&nbsp; Tổng ${totalMinutes.toLocaleString("vi-VN")} phút (${(totalMinutes / 90).toFixed(2)} ca tc)</p>
        </div>
        <button id="closeNurseDailyDetailModalBtn" style="border:none;background:#f1f5f9;cursor:pointer;border-radius:8px;padding:7px 16px;font-size:0.9rem;font-weight:600;white-space:nowrap;">Đóng</button>
      </div>
      <div style="overflow-x:auto;">
        <table style="width:100%;border-collapse:collapse;font-size:0.88rem;">
          <thead>
            <tr style="background:#f8fafc;">
              <th style="padding:8px 10px;text-align:center;border-bottom:2px solid #e2e8f0;min-width:40px;">STT</th>
              <th style="padding:8px 10px;text-align:center;border-bottom:2px solid #e2e8f0;min-width:90px;">Ngày</th>
              <th style="padding:8px 10px;text-align:center;border-bottom:2px solid #e2e8f0;min-width:60px;">Giờ</th>
              <th style="padding:8px 10px;border-bottom:2px solid #e2e8f0;min-width:160px;">Khách hàng</th>
              <th style="padding:8px 10px;border-bottom:2px solid #e2e8f0;min-width:160px;">Dịch vụ</th>
              <th style="padding:8px 10px;border-bottom:2px solid #e2e8f0;min-width:120px;">Loại ca</th>
              <th style="padding:8px 10px;text-align:center;border-bottom:2px solid #e2e8f0;min-width:70px;">Số phút</th>
            </tr>
          </thead>
          <tbody>${tbody}</tbody>
        </table>
      </div>
    </div>
  `;

  document.body.appendChild(modalEl);

  modalEl.addEventListener("click", (e) => {
    if (e.target.id === "closeNurseDailyDetailModalBtn" || e.target === modalEl) {
      modalEl.remove();
    }
  });
}

function renderNurseReportMatrix() {
  if (!els.reportsTable) return;
  const matrix = getNurseDetailedReportMatrix(reportFilterState.start, reportFilterState.end);
  const bodyRows = matrix.rows.length
    ? matrix.rows.map((row, index) => `
      <tr>
        <td style="text-align:center;">${index + 1}</td>
        <td style="font-weight:600;cursor:pointer;color:var(--primary,#5b5bd6);text-decoration:underline;" data-nurse-name="${encodeURIComponent(row.nurseName)}">${row.nurseName}</td>
        <td style="text-align:center;">${row.workingDays.toLocaleString("vi-VN")}</td>
        <td style="text-align:center;">${row.totalMinutes.toLocaleString("vi-VN")}</td>
        <td style="text-align:center;">${row.standardShiftCount.toFixed(2)}</td>
        <td style="text-align:center;">${row.travelAllowance.toLocaleString("vi-VN")} đ</td>
        ${NURSE_REPORT_BUCKETS.map((bucket) => `
          <td style="text-align:center;">${row.counts[bucket.key] || ""}</td>
        `).join("")}
      </tr>
    `).join("")
    : `<tr><td colspan="13" style="text-align:center;">Không có ca hoàn tất trong khoảng ngày đã chọn.</td></tr>`;

  els.reportsTable.innerHTML = `
    <thead>
      <tr>
        <th rowspan="2" style="min-width:70px;">STT</th>
        <th rowspan="2" style="min-width:220px;cursor:pointer;user-select:none;" data-nurse-sort="nurseName">Họ Tên${getNurseReportSortIndicator("nurseName")}</th>
        <th rowspan="2" style="cursor:pointer;user-select:none;" data-nurse-sort="workingDays">Ngày công${getNurseReportSortIndicator("workingDays")}</th>
        <th rowspan="2" style="cursor:pointer;user-select:none;" data-nurse-sort="totalMinutes">Tổng phút${getNurseReportSortIndicator("totalMinutes")}</th>
        <th rowspan="2" style="cursor:pointer;user-select:none;" data-nurse-sort="standardShiftCount">Ca tiêu chuẩn${getNurseReportSortIndicator("standardShiftCount")}</th>
        <th rowspan="2" style="cursor:pointer;user-select:none;" data-nurse-sort="travelAllowance">PP di chuyển${getNurseReportSortIndicator("travelAllowance")}</th>
        <th colspan="3">Tắm bé - Hợp đồng</th>
        <th colspan="4">Chăm sóc mẹ - Hợp đồng</th>
      </tr>
      <tr>
        <th style="cursor:pointer;user-select:none;" data-nurse-sort="baby30">Tắm bé 30P${getNurseReportSortIndicator("baby30")}</th>
        <th style="cursor:pointer;user-select:none;" data-nurse-sort="baby45">Tắm bé 45P${getNurseReportSortIndicator("baby45")}</th>
        <th style="cursor:pointer;user-select:none;" data-nurse-sort="baby60">Tắm bé 60P${getNurseReportSortIndicator("baby60")}</th>
        <th style="cursor:pointer;user-select:none;" data-nurse-sort="mother60">Mẹ 60P${getNurseReportSortIndicator("mother60")}</th>
        <th style="cursor:pointer;user-select:none;" data-nurse-sort="mother75">Ca mẹ 75P${getNurseReportSortIndicator("mother75")}</th>
        <th style="cursor:pointer;user-select:none;" data-nurse-sort="mother90">Mẹ 90P${getNurseReportSortIndicator("mother90")}</th>
        <th style="cursor:pointer;user-select:none;" data-nurse-sort="mother120">Mẹ 120P${getNurseReportSortIndicator("mother120")}</th>
      </tr>
    </thead>
    <tbody>
      <tr style="background:#faf5ff;font-weight:600;">
        <td style="text-align:center;">a1</td>
        <td>Định mức phút/ca</td>
        <td style="text-align:center;">Ngày có ca</td>
        <td style="text-align:center;">Tổng phút</td>
        <td style="text-align:center;">90p = 1 ca</td>
        <td style="text-align:center;">Theo km</td>
        <td style="text-align:center;">30</td>
        <td style="text-align:center;">45</td>
        <td style="text-align:center;">60</td>
        <td style="text-align:center;">60</td>
        <td style="text-align:center;">75</td>
        <td style="text-align:center;">90</td>
        <td style="text-align:center;">120</td>
      </tr>
      <tr style="background:#f8fafc;font-weight:700;">
        <td></td>
        <td>Số buổi</td>
        <td style="text-align:center;">${matrix.totalWorkingDays.toLocaleString("vi-VN")}</td>
        <td style="text-align:center;">${matrix.totalWorkedMinutes.toLocaleString("vi-VN")}</td>
        <td style="text-align:center;">${matrix.totalStandardShiftCount.toFixed(2)}</td>
        <td style="text-align:center;">${matrix.totalTravelAllowance.toLocaleString("vi-VN")} đ</td>
        ${NURSE_REPORT_BUCKETS.map((bucket) => `<td style="text-align:center;">${matrix.totals[bucket.key] ? matrix.totals[bucket.key].toLocaleString("vi-VN") : "-"}</td>`).join("")}
      </tr>
      ${bodyRows}
    </tbody>
  `;

  const nurseRowsInRange = getRowsByReportDepartment("nurse", reportFilterState.start, reportFilterState.end)
    .filter((item) => item.status === "completed");
  const telegramRowsInRange = nurseRowsInRange.filter((item) => {
    const source = String(item.source || "").toLowerCase();
    const id = String(item.id || "");
    return source.includes("telegram") || id.startsWith("tg-");
  });
  const yenRowsInRange = nurseRowsInRange.filter((item) => getCanonicalNurseName(item.nurse) === "Yến");
  const rowsWithDistance = nurseRowsInRange.filter((item) => getDistanceKm(item) > 0).length;
  const rowsWithTravelAllowance = nurseRowsInRange.filter((item) => getDistanceKm(item) > 10).length;

  els.reportsSummary.textContent = `Bộ lọc: ${reportFilterState.start} → ${reportFilterState.end} | Điều dưỡng | ${matrix.rows.length} điều dưỡng | ${matrix.overallTotal.toLocaleString("vi-VN")} ca hoàn tất | PP di chuyển: ${matrix.totalTravelAllowance.toLocaleString("vi-VN")} đ | Ca có km: ${rowsWithDistance} | Ca >10km: ${rowsWithTravelAllowance} | Telegram: ${telegramRowsInRange.length} ca | Yến: ${yenRowsInRange.length} ca`;
}

function getMarketingOwnerName(item) {
  const candidates = [item?.marketingStaff, item?.marketingName, item?.marketer, item?.ownerName];
  const matched = candidates.map((value) => String(value || "").trim()).find(Boolean);
  if (matched) return matched;
  return "Chưa gán";
}

function getMarketingBudget(item) {
  const candidates = [
    item?.marketingBudget,
    item?.budget,
    item?.adsBudget,
    item?.adSpend,
    item?.marketingCost,
    item?.campaignCost
  ];
  for (const value of candidates) {
    const parsed = parseFlexibleNumber(value);
    if (parsed > 0) return parsed;
  }
  return 0;
}

function getMarketingReportRows(start, end) {
  const rows = getRowsByReportDepartment("marketing", start, end);
  const bucket = new Map();

  rows.forEach((item) => {
    const date = item.registrationDate || "";
    const marketingName = getMarketingOwnerName(item);
    const key = `${date}||${marketingName}`;
    if (!bucket.has(key)) {
      bucket.set(key, {
        date,
        marketingName,
        budget: 0,
        messCount: 0,
        phones: new Set(),
        bookedCount: 0,
        contractCount: 0,
        revenue: 0
      });
    }

    const row = bucket.get(key);
    row.budget += getMarketingBudget(item);
    row.messCount += 1;
    const phone = String(item.phone || "").trim();
    if (phone) row.phones.add(phone);

    const isBooked = item.status === "confirmed" || item.status === "completed";
    if (isBooked) row.bookedCount += 1;

    const contractAmount = Number(item.contractAmount) || 0;
    if (contractAmount > 0) {
      row.contractCount += 1;
      row.revenue += contractAmount;
    }
  });

  let result = Array.from(bucket.values())
    .map((row) => {
      const phoneCount = row.phones.size;
      const costPerMess = row.messCount ? row.budget / row.messCount : 0;
      const costPerPhone = phoneCount ? row.budget / phoneCount : 0;
      const costPerBooked = row.bookedCount ? row.budget / row.bookedCount : 0;
      const costRate = row.revenue ? (row.budget / row.revenue) * 100 : 0;
      return {
        date: row.date,
        marketingName: row.marketingName,
        budget: row.budget,
        messCount: row.messCount,
        costPerMess,
        phoneCount,
        costPerPhone,
        bookedCount: row.bookedCount,
        costPerBooked,
        contractCount: row.contractCount,
        revenue: row.revenue,
        costRate
      };
    })
    .sort((a, b) => a.date.localeCompare(b.date) || a.marketingName.localeCompare(b.marketingName));

  if (marketingReportState.marketing) {
    result = result.filter((row) => row.marketingName === marketingReportState.marketing);
  }

  return result;
}

function renderMarketingReportTable(rows) {
  if (!els.reportsTable) return;

  const tbody = rows.length
    ? rows.map((row) => `
      <tr>
        <td>${row.date}</td>
        <td>${row.marketingName}</td>
        <td style="text-align:right;">${row.budget.toLocaleString("vi-VN")} đ</td>
        <td style="text-align:center;">${row.messCount.toLocaleString("vi-VN")}</td>
        <td style="text-align:right;">${row.costPerMess.toLocaleString("vi-VN")} đ</td>
        <td style="text-align:center;">${row.phoneCount.toLocaleString("vi-VN")}</td>
        <td style="text-align:right;">${row.costPerPhone.toLocaleString("vi-VN")} đ</td>
        <td style="text-align:center;">${row.bookedCount.toLocaleString("vi-VN")}</td>
        <td style="text-align:right;">${row.costPerBooked.toLocaleString("vi-VN")} đ</td>
        <td style="text-align:center;">${row.contractCount.toLocaleString("vi-VN")}</td>
        <td style="text-align:right;">${row.revenue.toLocaleString("vi-VN")} đ</td>
        <td style="text-align:center;">${row.costRate.toFixed(1)}%</td>
      </tr>
    `).join("")
    : '<tr><td colspan="12" style="text-align:center;">Không có dữ liệu marketing trong khoảng ngày đã chọn.</td></tr>';

  els.reportsTable.innerHTML = `
    <thead>
      <tr>
        <th>Ngày</th>
        <th>Tên marketing</th>
        <th style="text-align:right;">Ngân sách</th>
        <th style="text-align:center;">Lượng mess</th>
        <th style="text-align:right;">Chi phí/mess</th>
        <th style="text-align:center;">Số điện thoại</th>
        <th style="text-align:right;">Chi phí/SĐT</th>
        <th style="text-align:center;">Lịch trải nghiệm</th>
        <th style="text-align:right;">Chi phí/lịch</th>
        <th style="text-align:center;">Hợp đồng</th>
        <th style="text-align:right;">Doanh số</th>
        <th style="text-align:center;">% chi phí/doanh số</th>
      </tr>
    </thead>
    <tbody>${tbody}</tbody>
  `;
}

function renderStandardReportTable(detailRows) {
  if (!els.reportsTable) return;
  els.reportsTable.innerHTML = `
    <thead>
      <tr>
        <th id="reportsPrimaryCol">Ngày</th>
        <th id="reportsMetricCol1">Chỉ số 1</th>
        <th id="reportsMetricCol2">Chỉ số 2</th>
        <th id="reportsMetricCol3">Chỉ số 3</th>
        <th id="reportsMetricCol4">Chỉ số 4</th>
      </tr>
    </thead>
    <tbody id="reportsBody">${!detailRows.length
      ? '<tr><td colspan="5" style="text-align:center;">Không có dữ liệu trong khoảng ngày đã chọn.</td></tr>'
      : detailRows.map((row) => `<tr><td>${row.date}</td><td>${row.c1}</td><td>${row.c2}</td><td>${row.c3}</td><td>${row.c4}</td></tr>`).join("")}
    </tbody>
  `;

  els.reportsPrimaryCol = document.querySelector("#reportsPrimaryCol");
  els.reportsMetricCol1 = document.querySelector("#reportsMetricCol1");
  els.reportsMetricCol2 = document.querySelector("#reportsMetricCol2");
  els.reportsMetricCol3 = document.querySelector("#reportsMetricCol3");
  els.reportsMetricCol4 = document.querySelector("#reportsMetricCol4");
  els.reportsBody = document.querySelector("#reportsBody");
}

function toggleConsultantReportSort(sortKey) {
  if (consultantReportState.sortKey === sortKey) {
    consultantReportState.direction = consultantReportState.direction === "desc" ? "asc" : "desc";
    return;
  }
  consultantReportState.sortKey = sortKey;
  consultantReportState.direction = "desc";
}

function getConsultantSortIndicator(sortKey) {
  if (consultantReportState.sortKey !== sortKey) return "";
  return consultantReportState.direction === "desc" ? " ↓" : " ↑";
}

function getConsultantReportRows(start, end) {
  const rows = getRowsByReportDepartment("consultant", start, end);
  const bucket = new Map();

  rows.forEach((item) => {
    const date = String(item.registrationDate || today);
    const consultantName = String(item.consultant || "").trim();
    if (!consultantName) return;

    const key = `${date}__${consultantName}`;
    const prev = bucket.get(key) || {
      date,
      consultantName,
      receivedCount: 0,
      signedCount: 0,
      postponedCount: 0,
      revenue: 0,
      receivable: 0
    };

    const contractAmount = Number(item.contractAmount || 0);
    const status = String(item.status || "").toLowerCase();
    const noteText = normalizeTextForMatching(item.note || "");

    prev.receivedCount += 1;
    if (contractAmount > 0) {
      prev.signedCount += 1;
      prev.revenue += contractAmount;
      prev.receivable += contractAmount;
    }

    const isPostponed = status === "cancelled"
      || noteText.includes("hoan")
      || noteText.includes("doi lich")
      || noteText.includes("doi hen");
    if (isPostponed) prev.postponedCount += 1;

    bucket.set(key, prev);
  });

  const baseRows = Array.from(bucket.values()).map((row) => {
      const signRate = row.receivedCount ? (row.signedCount / row.receivedCount) * 100 : 0;
      const avgInvoice = row.signedCount ? row.revenue / row.signedCount : 0;
      return {
        ...row,
        signRate,
        avgInvoice
      };
    });

  const filteredRows = consultantReportState.consultant
    ? baseRows.filter((row) => row.consultantName === consultantReportState.consultant)
    : baseRows;

  const directionFactor = consultantReportState.direction === "asc" ? 1 : -1;
  return filteredRows.sort((a, b) => {
    const key = consultantReportState.sortKey;
    let compare = 0;
    if (key === "date") compare = a.date.localeCompare(b.date);
    else if (key === "consultantName") compare = a.consultantName.localeCompare(b.consultantName);
    else if (key === "receivedCount") compare = a.receivedCount - b.receivedCount;
    else if (key === "signedCount") compare = a.signedCount - b.signedCount;
    else if (key === "signRate") compare = a.signRate - b.signRate;
    else if (key === "postponedCount") compare = a.postponedCount - b.postponedCount;
    else if (key === "revenue") compare = a.revenue - b.revenue;
    else if (key === "receivable") compare = a.receivable - b.receivable;
    else if (key === "avgInvoice") compare = a.avgInvoice - b.avgInvoice;
    if (compare === 0) {
      return a.date.localeCompare(b.date) || a.consultantName.localeCompare(b.consultantName);
    }
    return compare * directionFactor;
  });
}

function renderConsultantReportTable(detailRows) {
  if (!els.reportsTable) return;
  const tbody = detailRows.length
    ? detailRows.map((row) => `
      <tr>
        <td>${row.date}</td>
        <td>${row.consultantName}</td>
        <td style="text-align:center;">${row.receivedCount.toLocaleString("vi-VN")}</td>
        <td style="text-align:center;">${row.signedCount.toLocaleString("vi-VN")}</td>
        <td style="text-align:center;">${row.signRate.toFixed(1)}%</td>
        <td style="text-align:center;">${row.postponedCount.toLocaleString("vi-VN")}</td>
        <td style="text-align:right;">${row.revenue.toLocaleString("vi-VN")} đ</td>
        <td style="text-align:right;">${row.receivable.toLocaleString("vi-VN")} đ</td>
        <td style="text-align:right;">${row.avgInvoice.toLocaleString("vi-VN")} đ</td>
      </tr>
    `).join("")
    : '<tr><td colspan="9" style="text-align:center;">Không có dữ liệu tư vấn trong khoảng ngày đã chọn.</td></tr>';

  els.reportsTable.innerHTML = `
    <thead>
      <tr>
        <th style="cursor:pointer;user-select:none;" data-consultant-sort="date">Ngày${getConsultantSortIndicator("date")}</th>
        <th style="cursor:pointer;user-select:none;" data-consultant-sort="consultantName">Tên tư vấn${getConsultantSortIndicator("consultantName")}</th>
        <th style="cursor:pointer;user-select:none;text-align:center;" data-consultant-sort="receivedCount">Số ca nhận${getConsultantSortIndicator("receivedCount")}</th>
        <th style="cursor:pointer;user-select:none;text-align:center;" data-consultant-sort="signedCount">Số ca ký${getConsultantSortIndicator("signedCount")}</th>
        <th style="cursor:pointer;user-select:none;text-align:center;" data-consultant-sort="signRate">Tỉ lệ ký${getConsultantSortIndicator("signRate")}</th>
        <th style="cursor:pointer;user-select:none;text-align:center;" data-consultant-sort="postponedCount">Số ca hoãn${getConsultantSortIndicator("postponedCount")}</th>
        <th style="cursor:pointer;user-select:none;text-align:right;" data-consultant-sort="revenue">Doanh số${getConsultantSortIndicator("revenue")}</th>
        <th style="cursor:pointer;user-select:none;text-align:right;" data-consultant-sort="receivable">Công nợ${getConsultantSortIndicator("receivable")}</th>
        <th style="cursor:pointer;user-select:none;text-align:right;" data-consultant-sort="avgInvoice">Đầu hoá đơn trung bình${getConsultantSortIndicator("avgInvoice")}</th>
      </tr>
    </thead>
    <tbody>${tbody}</tbody>
  `;
}

function toggleTelesaleReportSort(sortKey) {
  if (telesaleReportState.sortKey === sortKey) {
    telesaleReportState.direction = telesaleReportState.direction === "desc" ? "asc" : "desc";
    return;
  }
  telesaleReportState.sortKey = sortKey;
  telesaleReportState.direction = "desc";
}

function getTelesaleSortIndicator(sortKey) {
  if (telesaleReportState.sortKey !== sortKey) return "";
  return telesaleReportState.direction === "desc" ? " ↓" : " ↑";
}

function getTelesaleReportRows(start, end) {
  const rows = getRowsByReportDepartment("telesale", start, end);
  const bucket = new Map();
  rows.forEach((item) => {
    const saleName = String(item.saleStaff || "").trim();
    if (!saleName) return;
    const date = item.registrationDate || "";
    const key = `${date}||${saleName}`;
    if (!bucket.has(key)) {
      bucket.set(key, {
        date,
        saleName,
        items: []
      });
    }
    bucket.get(key).items.push(item);
  });

  let result = Array.from(bucket.values()).map((g) => {
    const messCount = g.items.length;
    const phones = new Set(g.items.map((item) => String(item.phone || "").trim()).filter(Boolean));
    const booked = g.items.filter((item) => item.status === "confirmed" || item.status === "completed");
    const cancelledOrDeferred = g.items.filter((item) => item.status === "cancelled" || item.status === "pending");
    const bookingRate = messCount ? (booked.length / messCount) * 100 : 0;
    const revenue = booked.reduce((sum, item) => sum + (Number(item.contractAmount) || 0), 0);
    return {
      date: g.date,
      saleName: g.saleName,
      messCount,
      phoneCount: phones.size,
      bookedCount: booked.length,
      cancelledCount: cancelledOrDeferred.length,
      bookingRate,
      revenue
    };
  });

  // filter by selected sale name
  if (telesaleReportState.sale) {
    result = result.filter((row) => row.saleName === telesaleReportState.sale);
  }

  // sort
  const dir = telesaleReportState.direction === "desc" ? -1 : 1;
  result.sort((a, b) => {
    const key = telesaleReportState.sortKey;
    const av = a[key]; const bv = b[key];
    if (typeof av === "string") return dir * av.localeCompare(bv);
    return dir * (av - bv);
  });

  return result;
}

function renderTelesaleReportTable(rows) {
  if (!els.reportsTable) return;
  const tbody = rows.length
    ? rows.map((row) => `
      <tr>
        <td>${row.date}</td>
        <td>${row.saleName}</td>
        <td style="text-align:center;">${row.messCount.toLocaleString("vi-VN")}</td>
        <td style="text-align:center;">${row.phoneCount.toLocaleString("vi-VN")}</td>
        <td style="text-align:center;">${row.bookedCount.toLocaleString("vi-VN")}</td>
        <td style="text-align:center;">${row.cancelledCount.toLocaleString("vi-VN")}</td>
        <td style="text-align:center;">${row.bookingRate.toFixed(1)}%</td>
        <td style="text-align:right;">${row.revenue.toLocaleString("vi-VN")} đ</td>
      </tr>
    `).join("")
    : '<tr><td colspan="8" style="text-align:center;">Không có dữ liệu telesale trong khoảng ngày đã chọn.</td></tr>';

  els.reportsTable.innerHTML = `
    <thead>
      <tr>
        <th style="cursor:pointer;user-select:none;" data-telesale-sort="date">Ngày${getTelesaleSortIndicator("date")}</th>
        <th style="cursor:pointer;user-select:none;" data-telesale-sort="saleName">Tên telesale${getTelesaleSortIndicator("saleName")}</th>
        <th style="cursor:pointer;user-select:none;text-align:center;" data-telesale-sort="messCount">Lượng mess${getTelesaleSortIndicator("messCount")}</th>
        <th style="cursor:pointer;user-select:none;text-align:center;" data-telesale-sort="phoneCount">Số điện thoại${getTelesaleSortIndicator("phoneCount")}</th>
        <th style="cursor:pointer;user-select:none;text-align:center;" data-telesale-sort="bookedCount">Lịch trải nghiệm${getTelesaleSortIndicator("bookedCount")}</th>
        <th style="cursor:pointer;user-select:none;text-align:center;" data-telesale-sort="cancelledCount">Ca hoãn huỷ${getTelesaleSortIndicator("cancelledCount")}</th>
        <th style="cursor:pointer;user-select:none;text-align:center;" data-telesale-sort="bookingRate">Tỉ lệ đặt lịch/mess${getTelesaleSortIndicator("bookingRate")}</th>
        <th style="cursor:pointer;user-select:none;text-align:right;" data-telesale-sort="revenue">Doanh số${getTelesaleSortIndicator("revenue")}</th>
      </tr>
    </thead>
    <tbody>${tbody}</tbody>
  `;
}

function getRowsByReportDepartment(departmentKey, start, end) {
  const inRange = schedules.filter((item) => {
    const d = item.registrationDate || "";
    if (start && d < start) return false;
    if (end && d > end) return false;
    return true;
  });
  if (departmentKey === "marketing") {
    const digital = ["facebook", "tiktok", "google", "website", "zalo", "instagram", "ads", "marketing"];
    return inRange.filter((item) => digital.some((key) => String(item.source || "").toLowerCase().includes(key)));
  }
  if (departmentKey === "telesale") return inRange.filter((item) => String(item.saleStaff || "").trim());
  if (departmentKey === "consultant") return inRange.filter((item) => String(item.consultant || "").trim());
  if (departmentKey === "nurse") return inRange.filter((item) => getCanonicalNurseName(item.nurse));
  return inRange;
}

function toDailyReportRows(departmentKey, rows) {
  if (departmentKey === "nurse") {
    const nurseBucket = new Map();
    rows.forEach((item) => {
      const nurseName = getCanonicalNurseName(item.nurse);
      if (!nurseName) return;
      const serviceText = String(item.service || "").trim();
      const serviceLower = serviceText.toLowerCase();
      const stageLower = String(item.stage || "").toLowerCase();
      const shiftMinutes = Number(item.shiftMinutes) > 0
        ? Number(item.shiftMinutes)
        : (() => {
            const matched = String(item.sessionDuration || "").match(/(\d{2,3})/);
            return matched ? Number(matched[1]) : 90;
          })();

      const prev = nurseBucket.get(nurseName) || {
        nurseName,
        totalShifts: 0,
        babyShifts: 0,
        motherShifts: 0,
        babyByDuration: new Map(),
        motherByDuration: new Map(),
        serviceCount: new Map(),
        otherShifts: 0
      };

      prev.totalShifts += 1;
      if (serviceText) {
        prev.serviceCount.set(serviceText, (prev.serviceCount.get(serviceText) || 0) + 1);
      }

      const isBabyCare = serviceLower.includes("bé") || stageLower.includes("em bé");
      const isMotherCare = !isBabyCare && (serviceLower.includes("mẹ") || stageLower.includes("mẹ") || stageLower.includes("bau"));

      if (isBabyCare) {
        prev.babyShifts += 1;
        prev.babyByDuration.set(shiftMinutes, (prev.babyByDuration.get(shiftMinutes) || 0) + 1);
      } else if (isMotherCare) {
        prev.motherShifts += 1;
        prev.motherByDuration.set(shiftMinutes, (prev.motherByDuration.get(shiftMinutes) || 0) + 1);
      } else {
        prev.otherShifts += 1;
      }

      nurseBucket.set(nurseName, prev);
    });

    const formatDurationDetail = (durationMap, preferredDurations) => {
      const ordered = [...preferredDurations, ...Array.from(durationMap.keys()).filter((d) => !preferredDurations.includes(d)).sort((a, b) => a - b)];
      const parts = ordered
        .filter((duration) => durationMap.get(duration))
        .map((duration) => `${duration}p:${durationMap.get(duration)}`);
      return parts.length ? parts.join(" | ") : "0 ca";
    };

    return Array.from(nurseBucket.values())
      .sort((a, b) => b.totalShifts - a.totalShifts || a.nurseName.localeCompare(b.nurseName))
      .map((row) => {
        const babyDetail = formatDurationDetail(row.babyByDuration, [30, 45, 60]);
        const motherDetail = formatDurationDetail(row.motherByDuration, [45, 60, 90, 120]);
        const topServices = Array.from(row.serviceCount.entries())
          .sort((a, b) => b[1] - a[1])
          .slice(0, 3)
          .map(([serviceName, count]) => `${serviceName} (${count})`)
          .join(", ");
        const extraNote = row.otherShifts ? ` | Ca khác: ${row.otherShifts}` : "";

        return {
          date: row.nurseName,
          c1: row.totalShifts.toLocaleString("vi-VN"),
          c2: `${row.babyShifts.toLocaleString("vi-VN")} ca (${babyDetail})`,
          c3: `${row.motherShifts.toLocaleString("vi-VN")} ca (${motherDetail})`,
          c4: `${topServices || "Chưa có dịch vụ cụ thể"}${extraNote}`
        };
      });
  }

  const bucket = new Map();
  rows.forEach((item) => {
    const key = item.registrationDate || today;
    const list = bucket.get(key) || [];
    list.push(item);
    bucket.set(key, list);
  });

  return Array.from(bucket.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([date, list]) => {
      if (departmentKey === "marketing") {
        const interactions = list.length;
        const phones = new Set(list.map((item) => String(item.phone || "").trim()).filter(Boolean)).size;
        const conversion = interactions ? (phones / interactions) * 100 : 0;
        const contract = list.reduce((sum, item) => sum + Number(item.contractAmount || 0), 0);
        return { date, c1: interactions.toLocaleString("vi-VN"), c2: phones.toLocaleString("vi-VN"), c3: `${conversion.toFixed(1)}%`, c4: `${contract.toLocaleString("vi-VN")} đ` };
      }
      if (departmentKey === "telesale") {
        const total = list.length;
        const booked = list.filter((item) => item.status === "confirmed" || item.status === "completed").length;
        const canceled = list.filter((item) => item.status === "cancelled").length;
        return { date, c1: total.toLocaleString("vi-VN"), c2: booked.toLocaleString("vi-VN"), c3: `${(total ? (booked / total) * 100 : 0).toFixed(1)}%`, c4: `${(total ? (canceled / total) * 100 : 0).toFixed(1)}%` };
      }
      if (departmentKey === "consultant") {
        const total = list.length;
        const signed = list.filter((item) => Number(item.contractAmount || 0) > 0);
        const revenue = signed.reduce((sum, item) => sum + Number(item.contractAmount || 0), 0);
        return { date, c1: total.toLocaleString("vi-VN"), c2: `${revenue.toLocaleString("vi-VN")} đ`, c3: signed.length.toLocaleString("vi-VN"), c4: `${(total ? (signed.length / total) * 100 : 0).toFixed(1)}%` };
      }
      const total = list.length;
      const caring = list.filter((item) => item.status === "pending" || item.status === "confirmed").length;
      const done = list.filter((item) => item.status === "completed").length;
      const score = total ? Math.min(5, (done / total) * 5) : 0;
      return { date, c1: caring.toLocaleString("vi-VN"), c2: done.toLocaleString("vi-VN"), c3: `${score.toFixed(1)}/5`, c4: `${(total ? (done / total) * 100 : 0).toFixed(1)}%` };
    });
}

function getActiveReportDailyRows() {
  if (!activeReportDepartment) return [];
  const rows = getRowsByReportDepartment(activeReportDepartment, reportFilterState.start, reportFilterState.end);
  return toDailyReportRows(activeReportDepartment, rows);
}

function exportReportDetailExcel() {
  if (!activeReportDepartment) {
    showToast("Vui lòng chọn một thư mục báo cáo trước khi xuất.", "warning");
    return;
  }
  const meta = REPORT_DEPARTMENT_META[activeReportDepartment];
  if (activeReportDepartment === "nurse") {
    const matrix = getNurseDetailedReportMatrix(reportFilterState.start, reportFilterState.end);
    if (!matrix.rows.length) {
      showToast("Không có dữ liệu để xuất Excel theo bộ lọc hiện tại.", "warning");
      return;
    }

    const header = ["STT", "Họ Tên", "Ngày công", "Tổng phút", "Ca tiêu chuẩn", ...NURSE_REPORT_BUCKETS.map((bucket) => bucket.label)];
    const durationRow = ["a1", "Định mức phút/ca", "Ngày có ca", "Tổng phút", "90p = 1 ca", ...NURSE_REPORT_BUCKETS.map((bucket) => bucket.minutes)];
    const totalRow = ["", "Số buổi", matrix.totalWorkingDays, matrix.totalWorkedMinutes, matrix.totalStandardShiftCount.toFixed(2), ...NURSE_REPORT_BUCKETS.map((bucket) => matrix.totals[bucket.key] || 0)];
    const records = matrix.rows.map((row, index) => [
      index + 1,
      row.nurseName,
      row.workingDays,
      row.totalMinutes,
      row.standardShiftCount.toFixed(2),
      ...NURSE_REPORT_BUCKETS.map((bucket) => row.counts[bucket.key] || 0)
    ]);
    const csv = [header, durationRow, totalRow, ...records]
      .map((line) => line.map((cell) => toCsvValue(cell)).join(","))
      .join("\n");
    const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8;" });
    downloadBlob(blob, `bao-cao-dieu-duong-${reportFilterState.start || "all"}-${reportFilterState.end || "all"}.csv`);
    showToast(`Đã xuất Excel báo cáo ${matrix.rows.length} điều dưỡng.`);
    logActivity("Báo cáo", "Xuất Excel", `${meta?.title || activeReportDepartment} | ${matrix.rows.length} điều dưỡng`);
    renderActivityTable();
    return;
  }

  if (activeReportDepartment === "marketing") {
    const marketingRows = getMarketingReportRows(reportFilterState.start, reportFilterState.end);
    if (!marketingRows.length) {
      showToast("Không có dữ liệu để xuất Excel theo bộ lọc hiện tại.", "warning");
      return;
    }
    const header = [
      "Ngày",
      "Tên marketing",
      "Ngân sách",
      "Lượng mess",
      "Chi phí/mess",
      "Số điện thoại",
      "Chi phí/SĐT",
      "Lịch trải nghiệm",
      "Chi phí/lịch",
      "Hợp đồng",
      "Doanh số",
      "% chi phí/doanh số"
    ];
    const records = marketingRows.map((row) => [
      row.date,
      row.marketingName,
      `${row.budget.toLocaleString("vi-VN")} đ`,
      row.messCount,
      `${row.costPerMess.toLocaleString("vi-VN")} đ`,
      row.phoneCount,
      `${row.costPerPhone.toLocaleString("vi-VN")} đ`,
      row.bookedCount,
      `${row.costPerBooked.toLocaleString("vi-VN")} đ`,
      row.contractCount,
      `${row.revenue.toLocaleString("vi-VN")} đ`,
      `${row.costRate.toFixed(1)}%`
    ]);
    const csv = [header, ...records].map((line) => line.map((cell) => toCsvValue(cell)).join(",")).join("\n");
    const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8;" });
    downloadBlob(blob, `bao-cao-marketing-${reportFilterState.start || "all"}-${reportFilterState.end || "all"}.csv`);
    showToast(`Đã xuất Excel báo cáo ${marketingRows.length} dòng.`);
    logActivity("Báo cáo", "Xuất Excel", `${meta?.title || activeReportDepartment} | ${marketingRows.length} dòng`);
    renderActivityTable();
    return;
  }

  if (activeReportDepartment === "consultant") {
    const consultantRows = getConsultantReportRows(reportFilterState.start, reportFilterState.end);
    if (!consultantRows.length) {
      showToast("Không có dữ liệu để xuất Excel theo bộ lọc hiện tại.", "warning");
      return;
    }
    const header = ["Ngày", "Tên tư vấn", "Số ca nhận", "Số ca ký", "Tỉ lệ ký", "Số ca hoãn", "Doanh số", "Công nợ", "Đầu hoá đơn trung bình"];
    const records = consultantRows.map((row) => [
      row.date,
      row.consultantName,
      row.receivedCount,
      row.signedCount,
      `${row.signRate.toFixed(1)}%`,
      row.postponedCount,
      `${row.revenue.toLocaleString("vi-VN")} đ`,
      `${row.receivable.toLocaleString("vi-VN")} đ`,
      `${row.avgInvoice.toLocaleString("vi-VN")} đ`
    ]);
    const csv = [header, ...records].map((line) => line.map((cell) => toCsvValue(cell)).join(",")).join("\n");
    const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8;" });
    downloadBlob(blob, `bao-cao-${activeReportDepartment}-${reportFilterState.start || "all"}-${reportFilterState.end || "all"}.csv`);
    showToast(`Đã xuất Excel báo cáo ${consultantRows.length} dòng.`);
    logActivity("Báo cáo", "Xuất Excel", `${meta?.title || activeReportDepartment} | ${consultantRows.length} dòng`);
    renderActivityTable();
    return;
  }

  if (activeReportDepartment === "telesale") {
    const telesaleRows = getTelesaleReportRows(reportFilterState.start, reportFilterState.end);
    if (!telesaleRows.length) {
      showToast("Không có dữ liệu để xuất Excel theo bộ lọc hiện tại.", "warning");
      return;
    }
    const header = ["Ngày", "Tên telesale", "Lượng mess", "Số điện thoại", "Lịch trải nghiệm", "Ca hoãn huỷ", "Tỉ lệ đặt lịch/mess", "Doanh số"];
    const records = telesaleRows.map((row) => [
      row.date,
      row.saleName,
      row.messCount,
      row.phoneCount,
      row.bookedCount,
      row.cancelledCount,
      `${row.bookingRate.toFixed(1)}%`,
      `${row.revenue.toLocaleString("vi-VN")} đ`
    ]);
    const csv = [header, ...records].map((line) => line.map((cell) => toCsvValue(cell)).join(",")).join("\n");
    const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8;" });
    downloadBlob(blob, `bao-cao-telesale-${reportFilterState.start || "all"}-${reportFilterState.end || "all"}.csv`);
    showToast(`Đã xuất Excel báo cáo ${telesaleRows.length} dòng.`);
    logActivity("Báo cáo", "Xuất Excel", `${meta?.title || activeReportDepartment} | ${telesaleRows.length} dòng`);
    renderActivityTable();
    return;
  }

  const rows = getActiveReportDailyRows();
  if (!rows.length) {
    showToast("Không có dữ liệu để xuất Excel theo bộ lọc hiện tại.", "warning");
    return;
  }

  const firstColumnLabel = activeReportDepartment === "nurse" ? "Điều dưỡng" : "Ngày";
  const header = [firstColumnLabel, ...(meta?.cols || ["Chỉ số 1", "Chỉ số 2", "Chỉ số 3", "Chỉ số 4"])];
  const records = rows.map((row) => [row.date, row.c1, row.c2, row.c3, row.c4]);
  const csv = [header, ...records].map((line) => line.map((cell) => toCsvValue(cell)).join(",")).join("\n");
  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8;" });
  downloadBlob(blob, `bao-cao-${activeReportDepartment}-${reportFilterState.start || "all"}-${reportFilterState.end || "all"}.csv`);
  showToast(`Đã xuất Excel báo cáo ${rows.length} dòng.`);
  logActivity("Báo cáo", "Xuất Excel", `${meta?.title || activeReportDepartment} | ${rows.length} dòng`);
  renderActivityTable();
}

async function exportReportDetailPdf() {
  if (!can("canExportPdf")) {
    showToast("Bạn không có quyền xuất PDF.", "warning");
    return;
  }
  if (!activeReportDepartment) {
    showToast("Vui lòng chọn một thư mục báo cáo trước khi xuất.", "warning");
    return;
  }
  const rows = activeReportDepartment === "consultant"
    ? getConsultantReportRows(reportFilterState.start, reportFilterState.end)
    : activeReportDepartment === "marketing"
    ? getMarketingReportRows(reportFilterState.start, reportFilterState.end)
    : activeReportDepartment === "telesale"
    ? getTelesaleReportRows(reportFilterState.start, reportFilterState.end)
    : getActiveReportDailyRows();
  if (!rows.length) {
    showToast("Không có dữ liệu để xuất PDF theo bộ lọc hiện tại.", "warning");
    return;
  }

  showToast("Đang tạo PDF báo cáo...", "info");
  const canvas = await html2canvas(els.reportsDetailView, { scale: 2, useCORS: true });
  const image = canvas.toDataURL("image/png");
  const pdf = new jsPDF("p", "mm", "a4");
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const imageHeight = (canvas.height * pageWidth) / canvas.width;

  let heightLeft = imageHeight;
  let position = 0;
  pdf.addImage(image, "PNG", 0, position, pageWidth, imageHeight);
  heightLeft -= pageHeight;

  while (heightLeft > 0) {
    position -= pageHeight;
    pdf.addPage();
    pdf.addImage(image, "PNG", 0, position, pageWidth, imageHeight);
    heightLeft -= pageHeight;
  }

  pdf.save(`bao-cao-${activeReportDepartment}-${today}.pdf`);
  showToast("Đã xuất PDF báo cáo thành công.");
  logActivity("Báo cáo", "Xuất PDF", `${activeReportDepartment} | ${rows.length} dòng`);
  renderActivityTable();
}

function renderReportsPage() {
  if (!els.reportsSection || !els.reportsFolderView || !els.reportsDetailView) return;
  const detailMeta = activeReportDepartment ? REPORT_DEPARTMENT_META[activeReportDepartment] : null;

  els.reportsFolderView.classList.toggle("hidden", Boolean(detailMeta));
  els.reportsDetailView.classList.toggle("hidden", !detailMeta);

  if (!detailMeta) return;

  els.reportsDetailTitle.textContent = detailMeta.title;
  els.reportsDetailDesc.textContent = detailMeta.desc;
  els.reportsStartDate.value = reportFilterState.start;
  els.reportsEndDate.value = reportFilterState.end;
  if (els.reportsMarketingFilterWrap) {
    els.reportsMarketingFilterWrap.classList.toggle("hidden", activeReportDepartment !== "marketing");
  }
  if (els.reportsConsultantFilterWrap) {
    els.reportsConsultantFilterWrap.classList.toggle("hidden", activeReportDepartment !== "consultant");
  }
  if (els.reportsTelesaleFilterWrap) {
    els.reportsTelesaleFilterWrap.classList.toggle("hidden", activeReportDepartment !== "telesale");
  }

  if (activeReportDepartment === "marketing" && els.reportsMarketingFilter) {
    const allMarketingRows = getRowsByReportDepartment("marketing", reportFilterState.start, reportFilterState.end);
    const marketingNames = Array.from(new Set(
      allMarketingRows.map((item) => getMarketingOwnerName(item)).filter(Boolean)
    )).sort((a, b) => a.localeCompare(b));

    if (marketingReportState.marketing && !marketingNames.includes(marketingReportState.marketing)) {
      marketingReportState.marketing = "";
    }

    els.reportsMarketingFilter.innerHTML = `<option value="">Tất cả marketing</option>${marketingNames
      .map((name) => `<option value="${name.replace(/"/g, "&quot;")}">${name}</option>`)
      .join("")}`;
    els.reportsMarketingFilter.value = marketingReportState.marketing;
  }

  if (activeReportDepartment === "telesale" && els.reportsTelesaleFilter) {
    const allTelesaleRows = getRowsByReportDepartment("telesale", reportFilterState.start, reportFilterState.end);
    const saleNames = Array.from(new Set(
      allTelesaleRows.map((item) => String(item.saleStaff || "").trim()).filter(Boolean)
    )).sort((a, b) => a.localeCompare(b));
    if (telesaleReportState.sale && !saleNames.includes(telesaleReportState.sale)) {
      telesaleReportState.sale = "";
    }
    els.reportsTelesaleFilter.innerHTML = `<option value="">Tất cả telesale</option>${saleNames
      .map((name) => `<option value="${name.replace(/"/g, "&quot;")}">${name}</option>`)
      .join("")}`;
    els.reportsTelesaleFilter.value = telesaleReportState.sale;
  }

  if (activeReportDepartment === "consultant" && els.reportsConsultantFilter) {
    const allConsultantRows = getRowsByReportDepartment("consultant", reportFilterState.start, reportFilterState.end);
    const consultantNames = Array.from(new Set(
      allConsultantRows
        .map((item) => String(item.consultant || "").trim())
        .filter(Boolean)
    )).sort((a, b) => a.localeCompare(b));

    if (consultantReportState.consultant && !consultantNames.includes(consultantReportState.consultant)) {
      consultantReportState.consultant = "";
    }

    els.reportsConsultantFilter.innerHTML = `<option value="">Tất cả tư vấn</option>${consultantNames
      .map((name) => `<option value="${name.replace(/"/g, "&quot;")}">${name}</option>`)
      .join("")}`;
    els.reportsConsultantFilter.value = consultantReportState.consultant;
  }

  if (activeReportDepartment === "nurse") {
    renderNurseReportMatrix();
    return;
  }

  if (activeReportDepartment === "marketing") {
    const marketingRows = getMarketingReportRows(reportFilterState.start, reportFilterState.end);
    renderMarketingReportTable(marketingRows);
    const totalBudget = marketingRows.reduce((sum, row) => sum + row.budget, 0);
    const totalMess = marketingRows.reduce((sum, row) => sum + row.messCount, 0);
    const totalPhones = marketingRows.reduce((sum, row) => sum + row.phoneCount, 0);
    const totalBooked = marketingRows.reduce((sum, row) => sum + row.bookedCount, 0);
    const totalContracts = marketingRows.reduce((sum, row) => sum + row.contractCount, 0);
    const totalRevenue = marketingRows.reduce((sum, row) => sum + row.revenue, 0);
    const costRate = totalRevenue ? (totalBudget / totalRevenue) * 100 : 0;
    els.reportsSummary.textContent = `Bộ lọc: ${reportFilterState.start} → ${reportFilterState.end} | Marketing | ${marketingRows.length} dòng | Ngân sách: ${totalBudget.toLocaleString("vi-VN")} đ | Mess: ${totalMess.toLocaleString("vi-VN")} | SĐT: ${totalPhones.toLocaleString("vi-VN")} | Lịch trải nghiệm: ${totalBooked.toLocaleString("vi-VN")} | Hợp đồng: ${totalContracts.toLocaleString("vi-VN")} | Doanh số: ${totalRevenue.toLocaleString("vi-VN")} đ | CP/DS: ${costRate.toFixed(1)}%`;
    return;
  }

  if (activeReportDepartment === "consultant") {
    const consultantRows = getConsultantReportRows(reportFilterState.start, reportFilterState.end);
    renderConsultantReportTable(consultantRows);
    const totalReceived = consultantRows.reduce((sum, row) => sum + row.receivedCount, 0);
    const totalSigned = consultantRows.reduce((sum, row) => sum + row.signedCount, 0);
    const totalPostponed = consultantRows.reduce((sum, row) => sum + row.postponedCount, 0);
    const totalRevenue = consultantRows.reduce((sum, row) => sum + row.revenue, 0);
    const totalReceivable = consultantRows.reduce((sum, row) => sum + row.receivable, 0);
    const signRate = totalReceived ? (totalSigned / totalReceived) * 100 : 0;
    els.reportsSummary.textContent = `Bộ lọc: ${reportFilterState.start} → ${reportFilterState.end} | Tư vấn | ${consultantRows.length} dòng | Ca nhận: ${totalReceived.toLocaleString("vi-VN")} | Ca ký: ${totalSigned.toLocaleString("vi-VN")} (${signRate.toFixed(1)}%) | Ca hoãn: ${totalPostponed.toLocaleString("vi-VN")} | Doanh số: ${totalRevenue.toLocaleString("vi-VN")} đ | Công nợ: ${totalReceivable.toLocaleString("vi-VN")} đ`;
    return;
  }

  if (activeReportDepartment === "telesale") {
    const telesaleRows = getTelesaleReportRows(reportFilterState.start, reportFilterState.end);
    renderTelesaleReportTable(telesaleRows);
    const totalMess = telesaleRows.reduce((sum, row) => sum + row.messCount, 0);
    const totalBooked = telesaleRows.reduce((sum, row) => sum + row.bookedCount, 0);
    const totalCancelled = telesaleRows.reduce((sum, row) => sum + row.cancelledCount, 0);
    const totalRevenue = telesaleRows.reduce((sum, row) => sum + row.revenue, 0);
    const overallRate = totalMess ? (totalBooked / totalMess) * 100 : 0;
    els.reportsSummary.textContent = `Bộ lọc: ${reportFilterState.start} → ${reportFilterState.end} | Telesale | ${telesaleRows.length} dòng | Mess: ${totalMess.toLocaleString("vi-VN")} | Lịch trải nghiệm: ${totalBooked.toLocaleString("vi-VN")} (${overallRate.toFixed(1)}%) | Ca hoãn huỷ: ${totalCancelled.toLocaleString("vi-VN")} | Doanh số: ${totalRevenue.toLocaleString("vi-VN")} đ`;
    return;
  }

  const detailRows = getActiveReportDailyRows();
  renderStandardReportTable(detailRows);

  if (els.reportsPrimaryCol) {
    els.reportsPrimaryCol.textContent = "Ngày";
  }
  if (els.reportsMetricCol1) els.reportsMetricCol1.textContent = detailMeta.cols[0];
  if (els.reportsMetricCol2) els.reportsMetricCol2.textContent = detailMeta.cols[1];
  if (els.reportsMetricCol3) els.reportsMetricCol3.textContent = detailMeta.cols[2];
  if (els.reportsMetricCol4) els.reportsMetricCol4.textContent = detailMeta.cols[3];

  const depMap = { marketing: "Marketing", telesale: "Telesale", consultant: "Tư vấn", nurse: "Điều dưỡng" };
  els.reportsSummary.textContent = `Bộ lọc: ${reportFilterState.start} → ${reportFilterState.end} | ${depMap[activeReportDepartment]} | ${detailRows.length} ngày dữ liệu`;
}

async function syncTelegramBeforeNurseReportRender() {
  if (!authState.loggedIn) return;
  const result = await runTelegramRealtimeSync(true, { fullSync: true });
  if (result?.error) {
    showToast(`Không thể tự đồng bộ Telegram: ${result.error}`, "warning");
  }
}

async function openReportDepartment(departmentKey) {
  if (!REPORT_DEPARTMENT_META[departmentKey]) return;
  if (departmentKey === "nurse") {
    // Avoid stale single-day filter causing missing monthly nurse shifts.
    if (reportFilterState.start && reportFilterState.start === reportFilterState.end) {
      reportFilterState = { start: `${today.slice(0, 7)}-01`, end: today };
      showToast("Đã tự đặt bộ lọc báo cáo điều dưỡng về phạm vi tháng hiện tại.", "info");
    }
    await syncTelegramBeforeNurseReportRender();
  }
  activeReportDepartment = departmentKey;
  renderReportsPage();
}

function closeReportDepartment() {
  activeReportDepartment = null;
  renderReportsPage();
}

function normalizeAccountingCashflowFilterState(rawState = {}) {
  let start = rawState.start || "";
  let end = rawState.end || "";
  if (start && end && start > end) {
    const temp = start;
    start = end;
    end = temp;
  }
  return { start, end };
}

function saveAccountingCashflowFilterState() {
  saveJSON(STORAGE.accountingCashflowFilters, accountingCashflowFilterState);
}

function getFilteredAccountingCashflowEntries() {
  const { start, end } = accountingCashflowFilterState;
  return accountingCashflowEntries
    .filter((entry) => {
      if (start && entry.date < start) return false;
      if (end && entry.date > end) return false;
      return true;
    })
    .sort((a, b) => b.date.localeCompare(a.date) || b.createdAt - a.createdAt);
}

function generateCashflowVoucherCode(type, dateValue) {
  const prefix = type === "expense" ? "PC" : "PT";
  const compactDate = String(dateValue || today).replaceAll("-", "").slice(2);
  const currentCount = accountingCashflowEntries.filter((entry) => entry.type === type && entry.date === dateValue).length + 1;
  return `${prefix}-${compactDate}-${String(currentCount).padStart(2, "0")}`;
}

function renderAccountingCashflowContent() {
  const rows = getFilteredAccountingCashflowEntries();
  const income = rows.filter((entry) => entry.type === "income").reduce((sum, entry) => sum + entry.amount, 0);
  const expense = rows.filter((entry) => entry.type === "expense").reduce((sum, entry) => sum + entry.amount, 0);
  const pending = rows.filter((entry) => entry.status === "pending").length;
  const tbody = rows.length
    ? rows.map((entry) => `
        <tr>
          <td>${entry.date}</td>
          <td>${entry.voucherCode}</td>
          <td>${entry.type === "income" ? "Phiếu thu" : "Phiếu chi"}</td>
          <td>${entry.category}</td>
          <td>${entry.counterparty}</td>
          <td>${entry.content}</td>
          <td>${entry.amount.toLocaleString("vi-VN")} đ</td>
          <td>${entry.method}</td>
          <td>${entry.creator}</td>
          <td>${entry.status === "approved" ? "Đã duyệt" : entry.status === "draft" ? "Nháp" : "Chờ duyệt"}</td>
        </tr>
      `).join("")
    : '<tr><td colspan="10" style="text-align:center;">Không có phiếu thu chi trong khoảng thời gian đã chọn.</td></tr>';

  return `
    <div class="reports-detail-head" style="margin-bottom:12px;">
      <div class="reports-detail-actions">
        <button class="btn secondary" type="button" id="openCashflowModalBtn">+ Tạo thu chi</button>
      </div>
    </div>
    <div class="reports-filter" style="margin-top:0;">
      <div>
        <label>Từ ngày</label>
        <input id="cashflowFilterStartDate" type="date" value="${accountingCashflowFilterState.start || ""}" />
      </div>
      <div>
        <label>Đến ngày</label>
        <input id="cashflowFilterEndDate" type="date" value="${accountingCashflowFilterState.end || ""}" />
      </div>
      <button class="btn secondary" type="button" id="applyCashflowFilterBtn">Áp dụng</button>
      <button class="btn secondary" type="button" id="resetCashflowFilterBtn">Đặt lại</button>
      <div class="muted">Bộ lọc: ${accountingCashflowFilterState.start || "--"} → ${accountingCashflowFilterState.end || "--"} | ${rows.length} phiếu</div>
    </div>
    <div class="kpis" style="margin-bottom:12px;">
      <article class="kpi"><div class="label">Tổng thu</div><div class="value">${income.toLocaleString("vi-VN")} đ</div></article>
      <article class="kpi"><div class="label">Tổng chi</div><div class="value">${expense.toLocaleString("vi-VN")} đ</div></article>
      <article class="kpi"><div class="label">Chênh lệch</div><div class="value">${(income - expense).toLocaleString("vi-VN")} đ</div></article>
      <article class="kpi"><div class="label">Phiếu chờ duyệt</div><div class="value">${pending}</div></article>
    </div>
    <div class="tables">
      <table>
        <thead><tr><th>Ngày</th><th>Mã phiếu</th><th>Loại</th><th>Khoản mục</th><th>Đối tượng</th><th>Nội dung</th><th>Số tiền</th><th>Phương thức</th><th>Người lập</th><th>Trạng thái</th></tr></thead>
        <tbody>${tbody}</tbody>
      </table>
    </div>
    <div class="customer-modal hidden" id="cashflowModal">
      <div class="customer-modal-backdrop" id="cashflowModalBackdrop"></div>
      <div class="customer-modal-panel card" style="max-width:780px;">
        <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;">
          <div>
            <h3 style="margin:0;">Tạo phiếu thu chi</h3>
            <p class="muted" style="margin:3px 0 0;">Nhập đầy đủ chứng từ để lưu vào sổ quỹ.</p>
          </div>
          <button class="btn warn" type="button" id="closeCashflowModalBtn">Đóng</button>
        </div>
        <div class="form-grid" style="grid-template-columns:repeat(2,minmax(0,1fr));margin-top:10px;">
          <div>
            <label>Loại phiếu</label>
            <select id="cashflowType">
              <option value="income">Phiếu thu</option>
              <option value="expense">Phiếu chi</option>
            </select>
          </div>
          <div>
            <label>Ngày hạch toán</label>
            <input id="cashflowDate" type="date" value="${today}" />
          </div>
          <div>
            <label>Khoản mục</label>
            <input id="cashflowCategory" placeholder="Ví dụ: Thu khách hàng, chi vật tư..." />
          </div>
          <div>
            <label>Đối tượng</label>
            <input id="cashflowCounterparty" placeholder="Khách hàng / Nhà cung cấp / Nhân sự" />
          </div>
          <div>
            <label>Số tiền</label>
            <input id="cashflowAmount" type="number" min="0" step="1000" placeholder="0" />
          </div>
          <div>
            <label>Phương thức thanh toán</label>
            <select id="cashflowMethod">
              <option>Tiền mặt</option>
              <option>Chuyển khoản</option>
              <option>Ví điện tử</option>
            </select>
          </div>
          <div>
            <label>Trạng thái</label>
            <select id="cashflowStatus">
              <option value="pending">Chờ duyệt</option>
              <option value="approved">Đã duyệt</option>
              <option value="draft">Nháp</option>
            </select>
          </div>
          <div>
            <label>Người lập phiếu</label>
            <input id="cashflowCreator" value="${getCurrentUser()?.fullName || getCurrentUser()?.username || "Hệ thống"}" />
          </div>
          <div style="grid-column:1/-1;">
            <label>Nội dung</label>
            <textarea id="cashflowContent" rows="3" placeholder="Mô tả chi tiết nội dung thu hoặc chi"></textarea>
          </div>
          <div style="grid-column:1/-1;display:flex;justify-content:flex-end;gap:8px;align-items:center;">
            <span class="alert" id="cashflowModalStatus" style="margin:0;flex:1;"></span>
            <button class="btn secondary" type="button" id="saveCashflowBtn">Lưu phiếu</button>
          </div>
        </div>
      </div>
    </div>
  `;
}

function parseClockTimeToMinutes(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;
  const match = raw.match(/^(\d{1,2}):(\d{2})/);
  if (!match) return null;
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (Number.isNaN(hour) || Number.isNaN(minute)) return null;
  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) return null;
  return hour * 60 + minute;
}

function normalizeHeaderKey(value) {
  return String(value || "").toLowerCase().trim();
}

const ATTENDANCE_VENDOR_PROFILES = {
  generic: {
    label: "Chuẩn chung",
    fields: {
      employeeCode: ["employeecode", "employeeid", "ma nhan vien", "mã nhân viên", "manv", "id"],
      employeeName: ["employeename", "name", "ho ten", "họ tên", "nhan su", "nhân sự"],
      department: ["department", "phong ban", "phòng ban", "bo phan", "bộ phận"],
      date: ["date", "ngay", "ngày", "workdate", "attendance_date"],
      checkin: ["checkin", "check_in", "in", "gio vao", "giờ vào", "time_in"],
      checkout: ["checkout", "check_out", "out", "gio ra", "giờ ra", "time_out"],
      workHours: ["workhours", "work_hours", "gio cong", "giờ công", "hours", "tong gio"],
      overtimeHours: ["overtime", "overtimehours", "ot", "ot_hours", "gio tang ca", "giờ tăng ca"],
      lateMinutes: ["lateminutes", "late", "di muon", "đi muộn", "late_minutes"]
    }
  },
  wiseeye: {
    label: "Wise Eye",
    fields: {
      employeeCode: ["ma cham cong", "mã chấm công", "ma nv", "mã nv", "id nhan vien", "id nhân viên"],
      employeeName: ["ten nhan vien", "tên nhân viên", "ho ten", "họ tên", "ten", "tên"],
      department: ["phong ban", "phòng ban", "bo phan", "bộ phận", "don vi", "đơn vị"],
      date: ["ngay", "ngày", "ngay cham cong", "ngày chấm công", "date"],
      checkin: ["gio vao", "giờ vào", "vao", "check in", "checkin", "gio vao 1", "giờ vào 1"],
      checkout: ["gio ra", "giờ ra", "ra", "check out", "checkout", "gio ra 1", "giờ ra 1"],
      workHours: ["gio cong", "giờ công", "tong gio", "tổng giờ", "cong", "công"],
      overtimeHours: ["gio tang ca", "giờ tăng ca", "tang ca", "tăng ca", "ot"],
      lateMinutes: ["di muon", "đi muộn", "phut di muon", "phút đi muộn", "late"]
    }
  }
};

function getAttendanceVendorProfile(vendor) {
  return ATTENDANCE_VENDOR_PROFILES[vendor] || ATTENDANCE_VENDOR_PROFILES.generic;
}

function normalizeAccountingAttendanceSource(raw = {}) {
  const sourceType = raw.type === "api" ? "api" : "sheet";
  const sourceVendor = raw.vendor === "wiseeye" ? "wiseeye" : "generic";
  const minuteCandidates = [5, 10, 15];
  const parsedMinutes = Number(raw.autoSyncMinutes);
  const autoSyncMinutes = minuteCandidates.includes(parsedMinutes) ? parsedMinutes : 10;
  return {
    type: sourceType,
    url: String(raw.url || "").trim(),
    vendor: sourceVendor,
    autoSyncEnabled: Boolean(raw.autoSyncEnabled),
    autoSyncMinutes,
    lastSyncedAt: Math.max(0, Number(raw.lastSyncedAt) || 0),
    lastWarning: String(raw.lastWarning || "").trim()
  };
}

function saveAccountingAttendanceSource() {
  accountingAttendanceSource = normalizeAccountingAttendanceSource(accountingAttendanceSource);
  saveJSON(STORAGE.accountingAttendanceSource, accountingAttendanceSource);
}

function firstAvailableValue(source, keys) {
  for (const key of keys) {
    const value = source[key];
    if (value !== undefined && value !== null && String(value).trim() !== "") return value;
  }
  return "";
}

function normalizeAttendanceRow(raw, vendor = "generic") {
  if (!raw || typeof raw !== "object") return null;
  const sourceObj = {};
  Object.keys(raw).forEach((key) => {
    sourceObj[normalizeHeaderKey(key)] = raw[key];
  });

  const profile = getAttendanceVendorProfile(vendor);
  const fallbackProfile = ATTENDANCE_VENDOR_PROFILES.generic;
  const keysOf = (field) => {
    const set = [...(profile.fields[field] || []), ...(fallbackProfile.fields[field] || [])].map((item) => normalizeHeaderKey(item));
    return Array.from(new Set(set));
  };

  const employeeName = String(firstAvailableValue(sourceObj, keysOf("employeeName"))).trim();
  if (!employeeName) return null;

  const employeeCode = String(firstAvailableValue(sourceObj, keysOf("employeeCode")) || "").trim();
  const department = String(firstAvailableValue(sourceObj, keysOf("department")) || "Chưa rõ").trim();

  const dateRaw = String(firstAvailableValue(sourceObj, keysOf("date")) || "").trim();
  const date = dateRaw.includes("-") ? dateRaw.slice(0, 10) : toDateKey(new Date(dateRaw || Date.now()));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return null;

  const checkin = String(firstAvailableValue(sourceObj, keysOf("checkin")) || "").trim();
  const checkout = String(firstAvailableValue(sourceObj, keysOf("checkout")) || "").trim();

  let workHours = Number(firstAvailableValue(sourceObj, keysOf("workHours")) || 0);
  if (Number.isNaN(workHours) || workHours < 0) workHours = 0;

  const checkinMinutes = parseClockTimeToMinutes(checkin);
  const checkoutMinutes = parseClockTimeToMinutes(checkout);
  if (workHours === 0 && checkinMinutes !== null && checkoutMinutes !== null && checkoutMinutes > checkinMinutes) {
    workHours = Number(((checkoutMinutes - checkinMinutes) / 60).toFixed(2));
  }

  let overtimeHours = Number(firstAvailableValue(sourceObj, keysOf("overtimeHours")) || 0);
  if (Number.isNaN(overtimeHours) || overtimeHours < 0) overtimeHours = 0;
  if (overtimeHours === 0 && workHours > 8) overtimeHours = Number((workHours - 8).toFixed(2));

  let lateMinutes = Number(firstAvailableValue(sourceObj, keysOf("lateMinutes")) || 0);
  if (Number.isNaN(lateMinutes) || lateMinutes < 0) lateMinutes = 0;
  if (lateMinutes === 0 && checkinMinutes !== null && checkinMinutes > 8 * 60) lateMinutes = checkinMinutes - 8 * 60;

  return {
    id: `at-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    employeeCode,
    employeeName,
    department,
    date,
    checkin,
    checkout,
    workHours,
    overtimeHours,
    lateMinutes
  };
}

function analyzeAttendanceColumns(importedRows, vendor = "generic") {
  const headerSet = new Set();
  importedRows.forEach((row) => {
    Object.keys(row || {}).forEach((key) => {
      headerSet.add(normalizeHeaderKey(key));
    });
  });
  const profile = getAttendanceVendorProfile(vendor);
  const fallbackProfile = ATTENDANCE_VENDOR_PROFILES.generic;
  const requiredFields = ["employeeName", "date"];
  const missingRequired = requiredFields.filter((field) => {
    const aliasSet = [...(profile.fields[field] || []), ...(fallbackProfile.fields[field] || [])].map((key) => normalizeHeaderKey(key));
    return !aliasSet.some((key) => headerSet.has(key));
  });
  return {
    headerCount: headerSet.size,
    missingRequired
  };
}

function normalizeAccountingAttendanceFilterState(rawState = {}) {
  let start = rawState.start || "";
  let end = rawState.end || "";
  if (start && end && start > end) {
    const temp = start;
    start = end;
    end = temp;
  }
  return { start, end };
}

function saveAccountingAttendanceFilterState() {
  saveJSON(STORAGE.accountingAttendanceFilters, accountingAttendanceFilterState);
}

function normalizeAccountingServicePayrollFilterState(rawState = {}) {
  let start = rawState.start || "";
  let end = rawState.end || "";
  if (start && end && start > end) {
    const temp = start;
    start = end;
    end = temp;
  }
  return { start, end };
}

function saveAccountingServicePayrollFilterState() {
  saveJSON(STORAGE.accountingServicePayrollFilters, accountingServicePayrollFilterState);
}

function getFilteredAttendanceEntries() {
  const { start, end } = accountingAttendanceFilterState;
  return accountingAttendanceEntries.filter((entry) => {
    if (start && entry.date < start) return false;
    if (end && entry.date > end) return false;
    return true;
  });
}

function getAttendanceSummaryRows() {
  const bucket = new Map();
  getFilteredAttendanceEntries().forEach((entry) => {
    const key = entry.employeeCode || entry.employeeName;
    const prev = bucket.get(key) || {
      employeeName: entry.employeeName,
      employeeCode: entry.employeeCode,
      department: entry.department,
      workHours: 0,
      overtimeHours: 0,
      lateMinutes: 0,
      workingDays: 0
    };
    prev.workHours += Number(entry.workHours) || 0;
    prev.overtimeHours += Number(entry.overtimeHours) || 0;
    prev.lateMinutes += Number(entry.lateMinutes) || 0;
    if ((Number(entry.workHours) || 0) >= 4) prev.workingDays += 1;
    bucket.set(key, prev);
  });
  return Array.from(bucket.values()).sort((a, b) => b.workingDays - a.workingDays || a.employeeName.localeCompare(b.employeeName));
}

function renderAccountingAttendanceContent() {
  const canSyncAttendance = can("canSyncData");
  const rows = getAttendanceSummaryRows();
  const lastSyncLabel = accountingAttendanceSource.lastSyncedAt
    ? new Date(accountingAttendanceSource.lastSyncedAt).toLocaleString("vi-VN", { hour12: false })
    : "Chưa đồng bộ";
  const enoughWork = rows.filter((row) => row.workingDays >= 24).length;
  const shortWork = rows.filter((row) => row.workingDays < 24).length;
  const overtimePeople = rows.filter((row) => row.overtimeHours > 0).length;
  const latePeople = rows.filter((row) => row.lateMinutes > 0).length;
  const tbody = rows.length
    ? rows.map((row) => `
        <tr>
          <td>${row.employeeCode || "--"}</td>
          <td>${row.employeeName}</td>
          <td>${row.department || "--"}</td>
          <td>${row.workingDays.toLocaleString("vi-VN")}</td>
          <td>${row.overtimeHours.toFixed(2)}</td>
          <td>${row.lateMinutes.toLocaleString("vi-VN")}</td>
          <td>${row.workingDays >= 24 ? "Đạt" : "Thiếu công"}</td>
        </tr>
      `).join("")
    : '<tr><td colspan="7" style="text-align:center;">Chưa có dữ liệu chấm công.</td></tr>';

  return `
    <section class="card section" style="margin-bottom:12px;">
      <h4 style="margin:0 0 8px;">Link dữ liệu máy chấm công</h4>
      <div class="form-grid" style="grid-template-columns:minmax(160px,220px) minmax(160px,220px) 1fr auto auto;align-items:end;">
        <div>
          <label>Hãng máy chấm công</label>
          <select id="attendanceVendorType">
            <option value="generic" ${accountingAttendanceSource.vendor === "generic" ? "selected" : ""}>Chuẩn chung</option>
            <option value="wiseeye" ${accountingAttendanceSource.vendor === "wiseeye" ? "selected" : ""}>Wise Eye</option>
          </select>
        </div>
        <div>
          <label>Loại nguồn</label>
          <select id="attendanceSourceType">
            <option value="sheet" ${accountingAttendanceSource.type === "sheet" ? "selected" : ""}>Google Sheets / CSV URL</option>
            <option value="api" ${accountingAttendanceSource.type === "api" ? "selected" : ""}>REST API JSON</option>
          </select>
        </div>
        <div>
          <label>URL xuất dữ liệu</label>
          <input id="attendanceSourceUrl" value="${accountingAttendanceSource.url || ""}" placeholder="https://..." />
        </div>
        <button class="btn secondary" type="button" id="syncAttendanceBtn" ${canSyncAttendance ? "" : "disabled"}>Đồng bộ chấm công</button>
        <button class="btn secondary" type="button" id="clearAttendanceBtn" ${canSyncAttendance ? "" : "disabled"}>Xóa dữ liệu</button>
      </div>
      <div class="form-grid" style="grid-template-columns:minmax(190px,260px) minmax(120px,160px) 1fr;margin-top:8px;align-items:end;">
        <label style="display:flex;gap:8px;align-items:center;margin:0;">
          <input id="attendanceAutoSyncEnabled" type="checkbox" ${accountingAttendanceSource.autoSyncEnabled ? "checked" : ""} ${canSyncAttendance ? "" : "disabled"} />
          <span>Tự động đồng bộ</span>
        </label>
        <div>
          <label>Mỗi (phút)</label>
          <select id="attendanceAutoSyncMinutes" ${canSyncAttendance ? "" : "disabled"}>
            <option value="5" ${Number(accountingAttendanceSource.autoSyncMinutes) === 5 ? "selected" : ""}>5</option>
            <option value="10" ${Number(accountingAttendanceSource.autoSyncMinutes) === 10 ? "selected" : ""}>10</option>
            <option value="15" ${Number(accountingAttendanceSource.autoSyncMinutes) === 15 ? "selected" : ""}>15</option>
          </select>
        </div>
        <div class="muted" style="display:flex;align-items:center;font-size:0.82rem;">Lần đồng bộ gần nhất: ${lastSyncLabel}</div>
      </div>
      ${accountingAttendanceSource.lastWarning ? `<div class="alert" style="margin-top:8px;">${accountingAttendanceSource.lastWarning}</div>` : ""}
      <p class="muted" style="margin:8px 0 0;font-size:0.82rem;">Hỗ trợ CSV/JSON. Cột gợi ý: employeeCode, employeeName, department, date, checkin, checkout, workHours, overtimeHours, lateMinutes.</p>
    </section>
    <div class="reports-filter" style="margin-top:0;margin-bottom:12px;">
      <div>
        <label>Từ ngày</label>
        <input id="attendanceFilterStartDate" type="date" value="${accountingAttendanceFilterState.start || ""}" />
      </div>
      <div>
        <label>Đến ngày</label>
        <input id="attendanceFilterEndDate" type="date" value="${accountingAttendanceFilterState.end || ""}" />
      </div>
      <button class="btn secondary" type="button" id="applyAttendanceFilterBtn">Áp dụng</button>
      <button class="btn secondary" type="button" id="resetAttendanceFilterBtn">Đặt lại</button>
      <div class="muted">Bộ lọc: ${accountingAttendanceFilterState.start || "--"} → ${accountingAttendanceFilterState.end || "--"} | ${rows.length} nhân sự</div>
    </div>
    <div class="kpis" style="margin-bottom:12px;">
      <article class="kpi"><div class="label">Nhân sự đủ công</div><div class="value">${enoughWork}</div></article>
      <article class="kpi"><div class="label">Thiếu công</div><div class="value">${shortWork}</div></article>
      <article class="kpi"><div class="label">Có tăng ca</div><div class="value">${overtimePeople}</div></article>
      <article class="kpi"><div class="label">Đi muộn</div><div class="value">${latePeople}</div></article>
    </div>
    <div class="tables">
      <table>
        <thead><tr><th>Mã NV</th><th>Nhân sự</th><th>Phòng ban</th><th>Ngày công</th><th>OT (giờ)</th><th>Đi muộn (phút)</th><th>Trạng thái</th></tr></thead>
        <tbody>${tbody}</tbody>
      </table>
    </div>
  `;
}

function getServiceShiftPayrollRows(start = accountingServicePayrollFilterState.start, end = accountingServicePayrollFilterState.end) {
  const grouped = new Map();
  schedules.forEach((item) => {
    if (!item || item.status !== "completed") return;
    const registrationDate = String(item.registrationDate || "");
    if (start && registrationDate < start) return;
    if (end && registrationDate > end) return;
    const nurseName = getCanonicalNurseName(item.nurse);
    if (!nurseName) return;

    const prev = grouped.get(nurseName) || {
      nurseName,
      shiftCount: 0,
      workingDates: new Set(),
      contractAmount: 0,
      minutesWorked: 0,
      travelAllowance: 0,
      renewContractAmount: 0,
      productSalesAmount: 0,
      hotBonus: 0,
      monthlyCompetitionBonus: 0,
      socialInsuranceDeduction: 0,
      trainingDeduction: 0,
      unionDeduction: 0,
      violationDeduction: 0
    };

    const distanceKm = getDistanceKm(item);
    let distanceAllowance = 0;
    if (distanceKm > 20) distanceAllowance = 20000;
    else if (distanceKm > 15) distanceAllowance = 15000;
    else if (distanceKm > 10) distanceAllowance = 10000;

    const shiftMinutes = Number(item.shiftMinutes) || 90;
    const rowHotBonus = Number(item.hotBonus) || 0;
    const rowRenewContractAmount = Number(item.renewContractAmount) || 0;
    const rowProductSalesAmount = Number(item.productSalesAmount) || 0;
    const rowMonthlyCompetitionBonus = Number(item.monthlyCompetitionBonus) || 0;
    const rowSocialInsuranceDeduction = Number(item.socialInsuranceDeduction) || 0;
    const rowTrainingDeduction = Number(item.trainingDeduction) || 0;
    const rowUnionDeduction = Number(item.unionDeduction) || 0;
    const rowViolationDeduction = Number(item.violationDeduction) || 0;

    prev.shiftCount += 1;
    prev.workingDates.add(item.registrationDate);
    prev.contractAmount += Number(item.contractAmount) || 0;
    prev.minutesWorked += shiftMinutes;
    prev.travelAllowance += distanceAllowance;
    prev.renewContractAmount += rowRenewContractAmount;
    prev.productSalesAmount += rowProductSalesAmount;
    prev.hotBonus += rowHotBonus;
    prev.monthlyCompetitionBonus = Math.max(prev.monthlyCompetitionBonus, rowMonthlyCompetitionBonus);
    prev.socialInsuranceDeduction = Math.max(prev.socialInsuranceDeduction, rowSocialInsuranceDeduction);
    prev.trainingDeduction = Math.max(prev.trainingDeduction, rowTrainingDeduction);
    prev.unionDeduction = Math.max(prev.unionDeduction, rowUnionDeduction);
    prev.violationDeduction = Math.max(prev.violationDeduction, rowViolationDeduction);
    grouped.set(nurseName, prev);
  });

  // Sync shiftCount, minutesWorked and travelAllowance from nurse report matrix (includes admin/CEO overrides)
  const matrixData = getNurseDetailedReportMatrix(start, end);
  matrixData.rows.forEach((matrixRow) => {
    const existing = grouped.get(matrixRow.nurseName);
    if (existing) {
      existing.shiftCount = matrixRow.standardShiftCount;
      existing.minutesWorked = matrixRow.totalMinutes;
      existing.travelAllowance = matrixRow.travelAllowance;
    } else if (matrixRow.standardShiftCount > 0) {
      grouped.set(matrixRow.nurseName, {
        nurseName: matrixRow.nurseName,
        shiftCount: matrixRow.standardShiftCount,
        workingDates: new Set(),
        contractAmount: 0,
        minutesWorked: matrixRow.totalMinutes,
        travelAllowance: matrixRow.travelAllowance,
        renewContractAmount: 0,
        productSalesAmount: 0,
        hotBonus: 0,
        monthlyCompetitionBonus: 0,
        socialInsuranceDeduction: 0,
        trainingDeduction: 0,
        unionDeduction: 0,
        violationDeduction: 0
      });
    }
  });

  return Array.from(grouped.values())
    .map((row) => ({
      nurseName: row.nurseName,
      shiftCount: row.shiftCount,
      workingDays: row.workingDates.size,
      contractAmount: row.contractAmount,
      minutesWorked: row.minutesWorked,
      travelAllowance: row.travelAllowance,
      renewContractAmount: row.renewContractAmount,
      productSalesAmount: row.productSalesAmount,
      hotBonus: row.hotBonus,
      monthlyCompetitionBonus: row.monthlyCompetitionBonus,
      socialInsuranceDeduction: row.socialInsuranceDeduction,
      trainingDeduction: row.trainingDeduction,
      unionDeduction: row.unionDeduction,
      violationDeduction: row.violationDeduction
    }))
    .sort((a, b) => b.shiftCount - a.shiftCount || a.nurseName.localeCompare(b.nurseName));
}

function renderAccountingServicePayrollContent() {
  const { start, end } = accountingServicePayrollFilterState;
  const payrollPeriodLabel = `${start || "--"} → ${end || "--"}`;
  const serviceRows = getServiceShiftPayrollRows(start, end);
  const servicePayrollFormulaHtml = `
    Kỳ lương đang xem: <strong>${payrollPeriodLabel}</strong><br />
    <strong>Số ca &amp; tổng phút</strong> được lấy trực tiếp từ Bảng báo cáo điều dưỡng, đồng bộ theo dữ liệu lịch hoàn tất và báo cáo Telegram.<br />
    Lương dịch vụ = Lương cơ bản 3.000.000 + (Tổng phút ca x Đơn giá phút theo mốc ca) + Phụ phí di chuyển + Thưởng nóng + 5% tái ký + Thưởng thi đua + Chiết khấu sản phẩm 10% - Khấu trừ.<br />
    5% tái ký của từng điều dưỡng = <strong>Doanh thu tái ký</strong> của chính điều dưỡng đó x 5%.<br />
    Hoa hồng bán sản phẩm của từng điều dưỡng = <strong>Doanh số bán sản phẩm</strong> của chính điều dưỡng đó x 10%.<br />
    Mốc đơn giá phút theo tổng ca: &lt;70: 1.000đ | 70-89: 1.150đ | 90-109: 1.300đ | 110-124: 1.400đ | ≥125: 1.500đ.<br />
    Phụ phí di chuyển theo từng ca: &gt;10km: 10.000đ | &gt;15km: 15.000đ | &gt;20km: 20.000đ.<br />
    Thi đua tự động theo tổng ca tháng (điều kiện &gt;90 ca): Top 1: 500.000đ | Top 2: 300.000đ | Top 3: 200.000đ.<br />
    Chỉ tính các lịch có trạng thái <strong>Hoàn tất</strong> và có điều dưỡng phụ trách.
  `;

  const BASE_SALARY = 3000000;
  const DEFAULT_SHIFT_MINUTES = 90;
  const RESIGN_BONUS_PERCENT = 0.05;
  const PRODUCT_COMMISSION_PERCENT = 0.1;
  const COMPETITION_MIN_SHIFTS = 90;
  const COMPETITION_TOP_BONUSES = [500000, 300000, 200000];

  const minuteRateByShiftTier = (shiftCount) => {
    if (shiftCount < 70) return 1000;
    if (shiftCount < 90) return 1150;
    if (shiftCount < 110) return 1300;
    if (shiftCount < 125) return 1400;
    return 1500;
  };

  const calculateHotBonusByContract = (contractAmount, customHotBonus = 0) => {
    if (customHotBonus > 0) return customHotBonus;
    if (contractAmount >= 150000000) return 2000000;
    if (contractAmount >= 100000000) return 1200000;
    if (contractAmount >= 70000000) return 700000;
    return 0;
  };

  const competitionCandidates = [...serviceRows]
    .filter((row) => Number(row.shiftCount) > COMPETITION_MIN_SHIFTS)
    .sort((a, b) => {
      const shiftDiff = Number(b.shiftCount) - Number(a.shiftCount);
      if (shiftDiff !== 0) return shiftDiff;
      const minuteDiff = Number(b.minutesWorked) - Number(a.minutesWorked);
      if (minuteDiff !== 0) return minuteDiff;
      return String(a.nurseName || "").localeCompare(String(b.nurseName || ""));
    })
    .slice(0, COMPETITION_TOP_BONUSES.length);

  const competitionBonusByNurse = new Map();
  competitionCandidates.forEach((row, index) => {
    competitionBonusByNurse.set(row.nurseName, COMPETITION_TOP_BONUSES[index] || 0);
  });

  const normalizedRows = serviceRows.map((row) => {
    const minuteRate = minuteRateByShiftTier(row.shiftCount);
    const minutesWorked = row.minutesWorked || row.shiftCount * DEFAULT_SHIFT_MINUTES;
    const shiftSalary = minutesWorked * minuteRate;
    const hotBonus = calculateHotBonusByContract(row.contractAmount, row.hotBonus);
    const renewRevenue = Number(row.renewContractAmount) || 0;
    const renewBonus = renewRevenue * RESIGN_BONUS_PERCENT;
    const productRevenue = Number(row.productSalesAmount) || 0;
    const productCommission = productRevenue * PRODUCT_COMMISSION_PERCENT;
    const monthlyCompetitionBonus = competitionBonusByNurse.has(row.nurseName)
      ? Number(competitionBonusByNurse.get(row.nurseName) || 0)
      : Number(row.monthlyCompetitionBonus || 0);
    const totalDeductions = row.socialInsuranceDeduction + row.trainingDeduction + row.unionDeduction + row.violationDeduction;
    const totalSalary = BASE_SALARY + shiftSalary + row.travelAllowance + hotBonus + renewBonus + monthlyCompetitionBonus + productCommission - totalDeductions;

    return {
      ...row,
      minuteRate,
      minutesWorked,
      baseSalary: BASE_SALARY,
      shiftSalary,
      hotBonus,
      renewRevenue,
      renewBonus,
      productRevenue,
      monthlyCompetitionBonus,
      productCommission,
      totalDeductions,
      totalSalary
    };
  });

  const totalShifts = normalizedRows.reduce((sum, row) => sum + row.shiftCount, 0);
  const totalMinutes = normalizedRows.reduce((sum, row) => sum + row.minutesWorked, 0);
  const totalServiceSalary = normalizedRows.reduce((sum, row) => sum + row.totalSalary, 0);
  const qualifiedHotBonusCount = normalizedRows.filter((row) => row.hotBonus > 0).length;

  const tbody = normalizedRows.length
    ? normalizedRows.map((row) => `
      <tr>
        <td>${row.nurseName}</td>
        <td>${row.workingDays.toLocaleString("vi-VN")}</td>
        <td>${Number(row.shiftCount).toFixed(2)}</td>
        <td>${row.minutesWorked.toLocaleString("vi-VN")}</td>
        <td>${row.minuteRate.toLocaleString("vi-VN")} đ/phút</td>
        <td>${row.baseSalary.toLocaleString("vi-VN")} đ</td>
        <td>${row.shiftSalary.toLocaleString("vi-VN")} đ</td>
        <td>${row.travelAllowance.toLocaleString("vi-VN")} đ</td>
        <td>${row.hotBonus.toLocaleString("vi-VN")} đ</td>
        <td>${row.renewRevenue.toLocaleString("vi-VN")} đ</td>
        <td>${row.renewBonus.toLocaleString("vi-VN")} đ</td>
        <td>${row.monthlyCompetitionBonus.toLocaleString("vi-VN")} đ</td>
        <td>${row.productRevenue.toLocaleString("vi-VN")} đ</td>
        <td>${row.productCommission.toLocaleString("vi-VN")} đ</td>
        <td>${row.totalDeductions.toLocaleString("vi-VN")} đ</td>
        <td><strong>${row.totalSalary.toLocaleString("vi-VN")} đ</strong></td>
      </tr>
    `).join("")
    : '<tr><td colspan="16" style="text-align:center;">Chưa có ca dịch vụ hoàn tất trong kỳ này.</td></tr>';

  return `
    <div id="servicePayrollInfoModal" class="service-payroll-info-modal hidden">
      <div id="servicePayrollInfoBackdrop" class="service-payroll-info-backdrop"></div>
      <div class="service-payroll-info-panel card">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;">
          <h4 style="margin:0;">Công thức tính lương dịch vụ theo ca thực tế</h4>
          <button class="btn secondary" type="button" id="closeServicePayrollInfoBtn">Đóng</button>
        </div>
        <div class="muted" style="line-height:1.6;margin-top:10px;">
          ${servicePayrollFormulaHtml}
        </div>
      </div>
    </div>
    <div class="reports-filter" style="margin-top:0;margin-bottom:12px;">
      <div>
        <label>Từ ngày</label>
        <input id="servicePayrollFilterStartDate" type="date" value="${start || ""}" />
      </div>
      <div>
        <label>Đến ngày</label>
        <input id="servicePayrollFilterEndDate" type="date" value="${end || ""}" />
      </div>
      <button class="btn secondary" type="button" id="applyServicePayrollFilterBtn">Áp dụng</button>
      <button class="btn secondary" type="button" id="resetServicePayrollFilterBtn">Đặt lại tháng này</button>
      <div class="muted">Bộ lọc: ${start || "--"} → ${end || "--"} | ${normalizedRows.length} điều dưỡng</div>
    </div>
    <div class="kpis service-payroll-kpis" style="margin-bottom:12px;">
      <article class="kpi"><div class="label">Tổng ca hoàn tất</div><div class="value">${totalShifts.toLocaleString("vi-VN")}</div></article>
      <article class="kpi"><div class="label">Tổng phút công</div><div class="value">${totalMinutes.toLocaleString("vi-VN")}</div></article>
      <article class="kpi"><div class="label">Nhân sự có ca</div><div class="value">${normalizedRows.length.toLocaleString("vi-VN")}</div></article>
      <article class="kpi"><div class="label">Có thưởng nóng</div><div class="value">${qualifiedHotBonusCount.toLocaleString("vi-VN")}</div></article>
      <article class="kpi"><div class="label">Tổng lương dịch vụ</div><div class="value">${totalServiceSalary.toLocaleString("vi-VN")} đ</div></article>
    </div>
    <div class="tables">
      <table>
        <thead><tr><th>Điều dưỡng</th><th>Ngày đi ca</th><th>Ca thực tế</th><th>Tổng phút</th><th>Đơn giá phút</th><th>Lương cơ bản</th><th>Lương ca</th><th>Di chuyển</th><th>Thưởng nóng</th><th>Doanh thu tái ký</th><th>5% tái ký</th><th>Thi đua</th><th>Doanh số bán SP</th><th>HH sản phẩm (10%)</th><th>Khấu trừ</th><th>Tổng lương</th></tr></thead>
        <tbody>${tbody}</tbody>
      </table>
    </div>
  `;
}

async function fetchAttendanceSourceRows(type, rawUrl) {
  if (!rawUrl) throw new Error("Thiếu URL nguồn dữ liệu chấm công.");
  const url = type === "sheet" ? normalizeSheetUrl(rawUrl) : rawUrl;
  const response = await fetch(url, { method: "GET" });
  if (!response.ok) throw new Error(`Không tải được dữ liệu (${response.status}).`);

  if (type === "sheet") {
    const text = await response.text();
    return parseCsvText(text);
  }

  const payload = await response.json();
  if (Array.isArray(payload)) return payload;
  if (payload && Array.isArray(payload.data)) return payload.data;
  throw new Error("Dữ liệu API phải là mảng hoặc object có trường data là mảng.");
}

function mergeAttendanceRows(importedRows, vendor = "generic") {
  const normalized = [];
  let invalidRows = 0;
  importedRows.forEach((row) => {
    const normalizedRow = normalizeAttendanceRow(row, vendor);
    if (!normalizedRow) {
      invalidRows += 1;
      return;
    }
    normalized.push(normalizedRow);
  });

  const columnAnalysis = analyzeAttendanceColumns(importedRows, vendor);
  const warnings = [];
  if (columnAnalysis.missingRequired.length) {
    warnings.push(`Thiếu cột bắt buộc: ${columnAnalysis.missingRequired.join(", ")}.`);
  }
  if (invalidRows > 0) {
    warnings.push(`${invalidRows} dòng bị bỏ qua do sai định dạng hoặc thiếu dữ liệu.`);
  }

  if (!normalized.length) {
    return {
      count: 0,
      warnings,
      invalidRows,
      missingRequired: columnAnalysis.missingRequired
    };
  }

  accountingAttendanceEntries = normalized;
  saveJSON(STORAGE.accountingAttendance, accountingAttendanceEntries);
  return {
    count: normalized.length,
    warnings,
    invalidRows,
    missingRequired: columnAnalysis.missingRequired
  };
}

async function runAttendanceSync(options = {}) {
  const { silent = false, fromAuto = false } = options;
  if (attendanceSyncInProgress) return;
  const sourceType = accountingAttendanceSource.type;
  const sourceUrl = accountingAttendanceSource.url;
  const sourceVendor = accountingAttendanceSource.vendor || "generic";
  if (!sourceUrl) {
    if (!silent) showToast("Vui lòng nhập URL nguồn dữ liệu chấm công.", "warning");
    return;
  }

  attendanceSyncInProgress = true;
  try {
    if (!silent) showToast("Đang đồng bộ dữ liệu chấm công...", "info");
    const importedRows = await fetchAttendanceSourceRows(sourceType, sourceUrl);
    const result = mergeAttendanceRows(importedRows, sourceVendor);
    if (!result.count) {
      const warnMessage = result.warnings.join(" ") || "Không có bản ghi chấm công hợp lệ.";
      accountingAttendanceSource.lastWarning = warnMessage;
      accountingAttendanceSource.lastSyncedAt = Date.now();
      saveAccountingAttendanceSource();
      renderAccountingPage();
      if (!silent) showToast(warnMessage, "warning");
      return;
    }

    accountingAttendanceSource.lastSyncedAt = Date.now();
    accountingAttendanceSource.lastWarning = result.warnings.join(" ");
    saveAccountingAttendanceSource();
    renderAccountingPage();

    const successMessage = result.warnings.length
      ? `Đồng bộ ${result.count} bản ghi. Cảnh báo: ${result.warnings.join(" ")}`
      : `Đồng bộ thành công ${result.count} bản ghi chấm công.`;
    if (!silent) showToast(successMessage, result.warnings.length ? "warning" : "success");

    if (!fromAuto || result.warnings.length) {
      logActivity("Kế toán", fromAuto ? "Tự động đồng bộ chấm công" : "Đồng bộ chấm công", `Nguồn: ${sourceType} | Vendor: ${sourceVendor} | Bản ghi: ${result.count}`);
      renderActivityTable();
    }
  } catch (error) {
    const errorMessage = `Lỗi đồng bộ chấm công: ${error.message}`;
    accountingAttendanceSource.lastSyncedAt = Date.now();
    accountingAttendanceSource.lastWarning = errorMessage;
    saveAccountingAttendanceSource();
    renderAccountingPage();
    if (!silent) showToast(errorMessage, "error");
  } finally {
    attendanceSyncInProgress = false;
  }
}

function configureAttendanceAutoSync() {
  const signature = [
    accountingAttendanceSource.autoSyncEnabled ? "1" : "0",
    accountingAttendanceSource.autoSyncMinutes,
    accountingAttendanceSource.type,
    accountingAttendanceSource.url,
    accountingAttendanceSource.vendor,
    authState.loggedIn ? "1" : "0",
    can("canSyncData") ? "1" : "0"
  ].join("|");

  if (attendanceAutoSyncSignature === signature) return;
  attendanceAutoSyncSignature = signature;

  if (attendanceAutoSyncTimer) {
    clearInterval(attendanceAutoSyncTimer);
    attendanceAutoSyncTimer = null;
  }

  if (!accountingAttendanceSource.autoSyncEnabled) return;
  if (!authState.loggedIn || !can("canSyncData") || !accountingAttendanceSource.url) return;

  const intervalMs = Number(accountingAttendanceSource.autoSyncMinutes) * 60 * 1000;
  attendanceAutoSyncTimer = setInterval(() => {
    if (!authState.loggedIn || !can("canSyncData")) return;
    runAttendanceSync({ silent: true, fromAuto: true });
  }, intervalMs);
}

const ACCOUNTING_FOLDER_META = {
  cashflow: {
    title: "Thu chi",
    desc: "Quản lý phiếu thu, phiếu chi và đối soát số dư quỹ theo ngày.",
    content: ""
  },
  "attendance-office": {
    title: "Tính công văn phòng",
    desc: "Quản lý công làm việc, ca trực, nghỉ phép và dữ liệu chấm công cho nhân sự văn phòng.",
    content: ""
  },
  "attendance-service": {
    title: "Tính lương điều dưỡng",
    desc: "Tính lương đội dịch vụ theo số ca thực tế đi hàng ngày.",
    content: ""
  },
  payroll: {
    title: "Tính lương",
    desc: "Mẫu bảng lương theo lương cơ bản, phụ cấp, thưởng KPI và khấu trừ.",
    content: `
      <div class="kpis" style="margin-bottom:12px;">
        <article class="kpi"><div class="label">Tổng quỹ lương</div><div class="value">386.400.000 đ</div></article>
        <article class="kpi"><div class="label">Phụ cấp</div><div class="value">28.600.000 đ</div></article>
        <article class="kpi"><div class="label">Thưởng KPI</div><div class="value">41.200.000 đ</div></article>
        <article class="kpi"><div class="label">Khấu trừ</div><div class="value">6.450.000 đ</div></article>
      </div>
      <section class="card section" style="margin-bottom:12px;">
        <h4 style="margin:0 0 8px;">Cấu phần mẫu</h4>
        <div class="reports-folder-grid" style="grid-template-columns:repeat(4,minmax(0,1fr));margin-top:0;">
          <article class="metrics-dept-card"><h4>Lương cơ bản</h4><div class="muted">Tính theo bậc và ngày công thực tế.</div></article>
          <article class="metrics-dept-card"><h4>Phụ cấp</h4><div class="muted">Ca trực, trách nhiệm, điện thoại.</div></article>
          <article class="metrics-dept-card"><h4>Thưởng</h4><div class="muted">KPI doanh số, chất lượng, tuân thủ.</div></article>
          <article class="metrics-dept-card"><h4>Khấu trừ</h4><div class="muted">BHXH, đi muộn, tạm ứng.</div></article>
        </div>
      </section>
      <div class="tables">
        <table>
          <thead><tr><th>Nhân sự</th><th>Lương cơ bản</th><th>Phụ cấp</th><th>Thưởng KPI</th><th>Khấu trừ</th><th>Thực lĩnh</th><th>Trạng thái</th></tr></thead>
          <tbody>
            <tr><td>Yến</td><td>12.000.000 đ</td><td>1.200.000 đ</td><td>2.500.000 đ</td><td>650.000 đ</td><td>15.050.000 đ</td><td>Chờ chốt</td></tr>
            <tr><td>Quyền</td><td>14.000.000 đ</td><td>1.500.000 đ</td><td>6.200.000 đ</td><td>1.100.000 đ</td><td>20.600.000 đ</td><td>Đã kiểm tra</td></tr>
            <tr><td>Hồ Trang</td><td>10.500.000 đ</td><td>800.000 đ</td><td>4.600.000 đ</td><td>450.000 đ</td><td>15.450.000 đ</td><td>Đã kiểm tra</td></tr>
          </tbody>
        </table>
      </div>
    `
  },
  "finance-report": {
    title: "Báo cáo tài chính",
    desc: "Mẫu tổng hợp doanh thu, chi phí, công nợ và kết quả kinh doanh theo kỳ.",
    content: `
      <div class="kpis" style="margin-bottom:12px;">
        <article class="kpi"><div class="label">Doanh thu tháng</div><div class="value">1.286.000.000 đ</div></article>
        <article class="kpi"><div class="label">Chi phí vận hành</div><div class="value">742.000.000 đ</div></article>
        <article class="kpi"><div class="label">Lợi nhuận gộp</div><div class="value">544.000.000 đ</div></article>
        <article class="kpi"><div class="label">Công nợ phải thu</div><div class="value">136.500.000 đ</div></article>
      </div>
      <section class="card section" style="margin-bottom:12px;">
        <h4 style="margin:0 0 8px;">Nhận định mẫu</h4>
        <div class="muted">Biên lợi nhuận đang giữ ổn định nhờ doanh thu gói liệu trình tăng và chi phí vật tư được kiểm soát dưới ngưỡng kế hoạch.</div>
      </section>
      <div class="tables">
        <table>
          <thead><tr><th>Chỉ tiêu</th><th>Kỳ này</th><th>Kỳ trước</th><th>Chênh lệch</th><th>Ghi chú</th></tr></thead>
          <tbody>
            <tr><td>Doanh thu thuần</td><td>1.286.000.000 đ</td><td>1.154.000.000 đ</td><td>+11.4%</td><td>Tăng ở nhóm gói mẹ và bé sau sinh</td></tr>
            <tr><td>Chi phí nhân sự</td><td>386.400.000 đ</td><td>372.500.000 đ</td><td>+3.7%</td><td>Tăng do bổ sung ca tăng cường</td></tr>
            <tr><td>Chi phí vật tư</td><td>142.800.000 đ</td><td>151.200.000 đ</td><td>-5.6%</td><td>Giảm nhờ tối ưu tồn kho</td></tr>
            <tr><td>Lợi nhuận trước thuế</td><td>468.200.000 đ</td><td>395.600.000 đ</td><td>+18.4%</td><td>Hiệu quả tốt hơn kế hoạch tháng</td></tr>
          </tbody>
        </table>
      </div>
    `
  }
};

function renderAccountingPage() {
  if (!els.accountingSection || !els.accountingFolderView || !els.accountingDetailView) return;
  const meta = activeAccountingFolder ? ACCOUNTING_FOLDER_META[activeAccountingFolder] : null;
  els.accountingFolderView.classList.toggle("hidden", Boolean(meta));
  els.accountingDetailView.classList.toggle("hidden", !meta);
  if (!meta) return;
  if (activeAccountingFolder === "attendance-service") {
    els.accountingDetailTitle.innerHTML = `${meta.title} <button class="service-payroll-info-trigger" id="openServicePayrollInfoBtn" type="button" title="Xem công thức tính lương" aria-label="Xem công thức tính lương">i</button>`;
  } else {
    els.accountingDetailTitle.textContent = meta.title;
  }
  els.accountingDetailDesc.textContent = meta.desc;
  if (activeAccountingFolder === "cashflow") {
    els.accountingDetailContent.innerHTML = renderAccountingCashflowContent();
    return;
  }
  if (activeAccountingFolder === "attendance-office") {
    els.accountingDetailContent.innerHTML = renderAccountingAttendanceContent();
    return;
  }
  if (activeAccountingFolder === "attendance-service") {
    els.accountingDetailContent.innerHTML = renderAccountingServicePayrollContent();
    return;
  }
  els.accountingDetailContent.innerHTML = meta.content;
}

function openAccountingFolder(folderKey) {
  if (!ACCOUNTING_FOLDER_META[folderKey]) return;
  activeAccountingFolder = folderKey;
  renderAccountingPage();
}

function closeAccountingFolder() {
  activeAccountingFolder = null;
  renderAccountingPage();
}

function drawChart(filteredReports) {
  const latest = getLatestRowsByDepartment(filteredReports);
  const labels = DEPARTMENTS;
  const points = DEPARTMENTS.map((dep) => latest.find((r) => r.department === dep)?.completion || 0);
  if (chart) chart.destroy();
  chart = new Chart(els.chart, {
    type: "bar",
    data: {
      labels,
      datasets: [{
        data: points,
        label: "% Hoàn thành",
        backgroundColor: ["#0f766e", "#14b8a6", "#0284c7", "#f59e0b", "#7c3aed", "#ef4444"],
        borderRadius: 8
      }]
    },
    options: {
      plugins: { legend: { display: false } },
      scales: { y: { beginAtZero: true, max: 100 } }
    }
  });
}

function renderTable(filteredReports) {
  const rows = [...filteredReports].sort((a, b) => b.updatedAt - a.updatedAt).slice(0, 25);
  els.reportBody.innerHTML = rows.map((r) => `
    <tr>
      <td>${r.date}</td>
      <td>${r.department}</td>
      <td>${r.completion}%</td>
      <td>${r.quality}</td>
      <td>${r.issues}</td>
      <td>${r.submitter}</td>
      <td>${new Date(r.updatedAt).toLocaleString("vi-VN", { hour12: false })}</td>
    </tr>
  `).join("");
}

function formatVnd(value) {
  return `${Math.round(value).toLocaleString("vi-VN")} đ`;
}

function getRevenueFromReport(row) {
  const monthlyTarget = DEPARTMENT_MONTHLY_REVENUE_TARGET[row.department] || 250000000;
  const dailyBase = monthlyTarget / 30;
  const performanceRatio = ((row.completion * 0.6) + (row.quality * 0.4)) / 100;
  const issuePenalty = Math.max(0.72, 1 - row.issues * 0.05);
  return dailyBase * performanceRatio * issuePenalty;
}

function getLatestRowsByDateDepartment(list) {
  const bucket = new Map();
  list.forEach((row) => {
    const key = `${row.date}|${row.department}`;
    const prev = bucket.get(key);
    if (!prev || prev.updatedAt < row.updatedAt) bucket.set(key, row);
  });
  return Array.from(bucket.values());
}

function drawHomeRevenueChart(filteredReports) {
  const rows = getLatestRowsByDateDepartment(filteredReports).sort((a, b) => a.date.localeCompare(b.date));
  const byDate = new Map();

  rows.forEach((row) => {
    const prev = byDate.get(row.date) || { revenue: 0, kpiSum: 0, count: 0 };
    prev.revenue += getRevenueFromReport(row);
    prev.kpiSum += ((row.completion * 0.6) + (row.quality * 0.4));
    prev.count += 1;
    byDate.set(row.date, prev);
  });

  const labels = Array.from(byDate.keys());
  const revenues = labels.map((dateKey) => byDate.get(dateKey).revenue);
  const kpiProgress = labels.map((dateKey) => {
    const item = byDate.get(dateKey);
    return item.count ? item.kpiSum / item.count : 0;
  });

  if (homeRevenueChart) homeRevenueChart.destroy();
  homeRevenueChart = new Chart(els.homeRevenueChart, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          type: "bar",
          label: "Doanh số",
          data: revenues,
          backgroundColor: "rgba(6, 182, 212, 0.45)",
          borderColor: "#0891b2",
          borderWidth: 1,
          yAxisID: "y"
        },
        {
          type: "line",
          label: "Tiến độ KPI",
          data: kpiProgress,
          borderColor: "#0f766e",
          backgroundColor: "rgba(15, 118, 110, 0.12)",
          tension: 0.3,
          pointRadius: 3,
          yAxisID: "y1"
        }
      ]
    },
    options: {
      plugins: { legend: { position: "top" } },
      scales: {
        y: { beginAtZero: true },
        y1: { beginAtZero: true, max: 100, position: "right", grid: { drawOnChartArea: false } }
      }
    }
  });
}

function renderHomeOverview(filteredReports) {
  const latestToday = getLatestRowsByDepartment(filteredReports);
  const latestReport = filteredReports.length ? [...filteredReports].sort((a, b) => b.updatedAt - a.updatedAt)[0] : null;
  const latestByDayDepartment = getLatestRowsByDateDepartment(filteredReports);
  const anchorDate = latestReport ? latestReport.date : filterState.end;
  const monthKey = anchorDate.slice(0, 7);
  const dailyRows = latestByDayDepartment.filter((row) => row.date === anchorDate);

  const todayRevenue = dailyRows.reduce((sum, row) => sum + getRevenueFromReport(row), 0);
  const monthRevenue = latestByDayDepartment
    .filter((row) => row.date.startsWith(monthKey))
    .reduce((sum, row) => sum + getRevenueFromReport(row), 0);
  const totalRevenue = latestByDayDepartment.reduce((sum, row) => sum + getRevenueFromReport(row), 0);

  const monthTarget = Object.values(DEPARTMENT_MONTHLY_REVENUE_TARGET).reduce((sum, target) => sum + target, 0);
  const monthProgress = monthTarget > 0 ? ((monthRevenue / monthTarget) * 100).toFixed(1) : "0.0";

  els.homeDailyRevenue.textContent = formatVnd(todayRevenue);
  els.homeDailyRevenueMeta.textContent = `Mốc ngày: ${anchorDate}`;
  els.homeMonthlyRevenue.textContent = formatVnd(monthRevenue);
  els.homeMonthlyRevenueMeta.textContent = `Tháng ${monthKey.slice(5, 7)}/${monthKey.slice(0, 4)} - Hoàn thành ${monthProgress}%`;
  els.homeTotalRevenue.textContent = formatVnd(totalRevenue);
  els.homeTotalRevenueMeta.textContent = `Dữ liệu trong kỳ: ${filterState.start} đến ${filterState.end}`;

  const departmentRows = DEPARTMENTS.map((department) => {
    const rowsByDepartment = latestByDayDepartment.filter((item) => item.department === department);
    const row = rowsByDepartment.sort((a, b) => b.updatedAt - a.updatedAt)[0];
    if (!row) return { department, score: 0, revenue: 0, risk: "Chưa có dữ liệu" };
    const score = ((row.completion * 0.55) + (row.quality * 0.35) - (row.issues * 3.5));
    const revenue = rowsByDepartment.reduce((sum, item) => sum + getRevenueFromReport(item), 0);
    const risk = row.issues >= 3 || row.completion < 70 ? "Cao" : row.issues >= 1 ? "Trung bình" : "Thấp";
    return { department, score, revenue, risk };
  });

  els.homeDeptBody.innerHTML = departmentRows
    .map((item) => `
      <tr>
        <td>${item.department}</td>
        <td>${item.score.toFixed(1)}</td>
        <td>${formatVnd(item.revenue)}</td>
        <td>${item.risk}</td>
      </tr>
    `)
    .join("");
}

function renderLastUpdated() {
  if (!els.lastUpdated) return;

  if (!can("canViewData")) {
    els.lastUpdated.textContent = "Cập nhật gần nhất: --";
    return;
  }

  const latest = reports.length ? Math.max(...reports.map((r) => r.updatedAt)) : Date.now();
  els.lastUpdated.textContent = `Cập nhật gần nhất: ${new Date(latest).toLocaleString("vi-VN", { hour12: false })}`;
}

function renderAll() {
  setAuthUI();
  renderNewsPage();
  renderUserTable();
  renderScheduleStaffControls();
  renderCustomerCarePage();
  renderMetricsFilterControls();
  renderAccountingPage();
  renderReportsPage();
  renderCustomerTable();
  renderInventoryTable();
  renderInventoryStatsAndHistory();
  renderWorkflowDetailView();
  renderPolicyFolders();
  renderDepartmentMetrics();
  renderActivityTable();
  const filteredReports = getFilteredReports();
  renderHomeOverview(filteredReports);

  if (authState.loggedIn && activePage === "home") {
    drawHomeRevenueChart(filteredReports);
  } else if (homeRevenueChart) {
    homeRevenueChart.destroy();
    homeRevenueChart = null;
  }

  if (can("canViewData")) {
    updateKPI(filteredReports);
    renderTable(filteredReports);
  }
  renderLastUpdated();
  configureAttendanceAutoSync();
}

els.brandBlock.addEventListener("click", () => {
  if (authState.loggedIn) {
    els.logoUpload.click();
  }
});

els.menuToggle.addEventListener("click", () => {
  if (els.appMenu.classList.contains("open")) {
    closeMenu();
    return;
  }
  openMenu();
});

els.menuOverlay.addEventListener("click", closeMenu);

els.openInventoryModalBtn.addEventListener("click", openInventoryModal);
els.openInventoryHistoryBtn.addEventListener("click", () => {
  renderInventoryStatsAndHistory();
  els.inventoryHistoryModal.classList.remove("hidden");
});
els.closeInventoryHistoryBtn.addEventListener("click", () => {
  els.inventoryHistoryModal.classList.add("hidden");
});
els.inventoryHistoryModalBackdrop.addEventListener("click", () => {
  els.inventoryHistoryModal.classList.add("hidden");
});
els.closeInventoryModalBtn.addEventListener("click", closeInventoryModal);
els.inventoryModalBackdrop.addEventListener("click", closeInventoryModal);
els.closeInventoryTxnModalBtn.addEventListener("click", closeInventoryTxnModal);
els.inventoryTxnModalBackdrop.addEventListener("click", closeInventoryTxnModal);
els.inventorySearch.addEventListener("input", () => {
  renderInventoryTable();
  renderScheduleTable();
  renderInventoryStatsAndHistory();
});
els.inventoryFilterStatus.addEventListener("change", () => {
  renderInventoryTable();
  renderInventoryStatsAndHistory();
});
els.inventoryFilterStock.addEventListener("change", () => {
  renderInventoryTable();
  renderInventoryStatsAndHistory();
});
els.applyInventoryStatsBtn.addEventListener("click", () => {
  if (!can("canSubmitReport")) return;
  if (els.inventoryStatsStart.value && els.inventoryStatsEnd.value && els.inventoryStatsStart.value > els.inventoryStatsEnd.value) {
    els.inventoryStatusMessage.textContent = "Khoảng thời gian không hợp lệ: Từ ngày phải nhỏ hơn hoặc bằng Đến ngày.";
    return;
  }
  els.inventoryStatusMessage.textContent = "";
  inventoryStatsState = {
    start: els.inventoryStatsStart.value || "",
    end: els.inventoryStatsEnd.value || ""
  };
  renderInventoryStatsAndHistory();
});
els.resetInventoryStatsBtn.addEventListener("click", () => {
  if (!can("canSubmitReport")) return;
  inventoryStatsState = { start: "", end: "" };
  els.inventoryStatsStart.value = "";
  els.inventoryStatsEnd.value = "";
  els.inventoryStatusMessage.textContent = "";
  renderInventoryStatsAndHistory();
});
els.exportInventoryExcelBtn.addEventListener("click", exportFilteredInventoryToExcel);
els.exportInventoryPdfBtn.addEventListener("click", async () => {
  await exportFilteredInventoryToPdf();
});

els.saveInventoryBtn.addEventListener("click", () => {
  if (!can("canSubmitReport")) {
    els.inventoryModalStatus.textContent = "Bạn không có quyền quản lý kho.";
    return;
  }

  const productCode = els.inventoryProductCode.value.trim().toUpperCase();
  const productName = els.inventoryProductName.value.trim();
  const purchasePrice = Number(els.inventoryPurchasePrice.value);
  const salePrice = Number(els.inventorySalePrice.value);
  const quantity = Number(els.inventoryQuantity.value);
  const alertThreshold = Number(els.inventoryAlertThreshold.value);
  const status = els.inventoryProductStatus.value;
  const supplier = els.inventorySupplier.value.trim();
  const expiryDate = els.inventoryExpiryDate.value;

  if (!productCode || !productName) {
    els.inventoryModalStatus.textContent = "Vui lòng nhập mã sản phẩm và tên sản phẩm.";
    return;
  }

  if ([purchasePrice, salePrice, quantity, alertThreshold].some((n) => Number.isNaN(n) || n < 0)) {
    els.inventoryModalStatus.textContent = "Giá, số lượng và ngưỡng cảnh báo phải là số không âm.";
    return;
  }

  const duplicated = inventoryItems.find((item) => item.productCode.toUpperCase() === productCode && item.id !== editingInventoryId);
  if (duplicated) {
    els.inventoryModalStatus.textContent = "Mã sản phẩm đã tồn tại.";
    return;
  }

  if (editingInventoryId) {
    inventoryItems = inventoryItems.map((item) => {
      if (item.id !== editingInventoryId) return item;
      return {
        ...item,
        productCode,
        productName,
        purchasePrice,
        salePrice,
        quantity,
        alertThreshold,
        status,
        supplier,
        expiryDate,
        updatedAt: Date.now()
      };
    });
    logActivity("Kho", "Cập nhật vật tư", `${productCode} | ${productName}`);
  } else {
    inventoryItems.unshift({
      id: `inv-${Date.now()}`,
      productCode,
      productName,
      purchasePrice,
      salePrice,
      quantity,
      alertThreshold,
      status,
      supplier,
      expiryDate,
      updatedAt: Date.now()
    });
    logActivity("Kho", "Thêm vật tư", `${productCode} | ${productName}`);
  }

  saveJSON(STORAGE.inventoryItems, inventoryItems);
  closeInventoryModal();
  showToast("Đã lưu thông tin vật tư.");
  renderAll();
});

els.inventoryBody.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  const itemId = target.dataset.itemId;
  if (!itemId) return;

  if (target.classList.contains("user-action-toggle")) {
    const menu = document.querySelector(`.inventory-action-menu[data-item-id="${itemId}"]`);
    if (menu instanceof HTMLElement) openActionMenuAtToggle(target, menu);
    return;
  }

  if (target.classList.contains("inventory-in-btn")) {
    if (!can("canSubmitReport")) return;
    hideAllActionMenus();
    openInventoryTxnModal(itemId, "in");
    return;
  }

  if (target.classList.contains("inventory-out-btn")) {
    if (!can("canSubmitReport")) return;
    hideAllActionMenus();
    openInventoryTxnModal(itemId, "out");
    return;
  }

  if (target.classList.contains("inventory-edit-btn")) {
    if (!can("canSubmitReport")) return;
    hideAllActionMenus();
    const item = inventoryItems.find((row) => row.id === itemId);
    if (!item) return;
    editingInventoryId = item.id;
    els.inventoryModalTitle.textContent = `Chỉnh sửa: ${item.productName}`;
    els.saveInventoryBtn.textContent = "Cập nhật vật tư";
    els.inventoryProductCode.value = item.productCode;
    els.inventoryProductName.value = item.productName;
    els.inventoryPurchasePrice.value = String(item.purchasePrice);
    els.inventorySalePrice.value = String(item.salePrice);
    els.inventoryQuantity.value = String(item.quantity);
    els.inventoryAlertThreshold.value = String(item.alertThreshold);
    els.inventorySupplier.value = item.supplier || "";
    els.inventoryExpiryDate.value = item.expiryDate || "";
    els.inventoryProductStatus.value = item.status || "active";
    els.inventoryModalStatus.textContent = "";
    els.inventoryModal.classList.remove("hidden");
    return;
  }

  if (target.classList.contains("inventory-delete-btn")) {
    if (!can("canSubmitReport")) return;
    hideAllActionMenus();
    const deletedItem = inventoryItems.find((row) => row.id === itemId);
    inventoryItems = inventoryItems.filter((row) => row.id !== itemId);
    saveJSON(STORAGE.inventoryItems, inventoryItems);
    if (deletedItem) {
      logActivity("Kho", "Xóa vật tư", `${deletedItem.productCode} | ${deletedItem.productName}`);
    }
    els.inventoryStatusMessage.textContent = "Đã xóa vật tư khỏi kho.";
    renderAll();
  }
});

els.saveInventoryTxnBtn.addEventListener("click", () => {
  if (!can("canSubmitReport")) {
    els.inventoryTxnStatus.textContent = "Bạn không có quyền thao tác xuất/nhập kho.";
    return;
  }

  if (!inventoryTxnItemId) return;
  const txnType = els.inventoryTxnType.value;
  const qty = Number(els.inventoryTxnQty.value);
  const note = els.inventoryTxnNote.value.trim();

  if (!qty || Number.isNaN(qty) || qty <= 0) {
    els.inventoryTxnStatus.textContent = "Số lượng giao dịch phải lớn hơn 0.";
    return;
  }

  const item = inventoryItems.find((row) => row.id === inventoryTxnItemId);
  if (!item) return;

  if (txnType === "out" && qty > item.quantity) {
    els.inventoryTxnStatus.textContent = `Không đủ tồn kho. Hiện còn ${item.quantity}.`;
    return;
  }

  const nextQty = txnType === "in" ? item.quantity + qty : item.quantity - qty;
  inventoryItems = inventoryItems.map((row) => {
    if (row.id !== inventoryTxnItemId) return row;
    return {
      ...row,
      quantity: nextQty,
      updatedAt: Date.now()
    };
  });

  saveJSON(STORAGE.inventoryItems, inventoryItems);
  inventoryTransactions.unshift({
    id: `txn-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    itemId: item.id,
    productCode: item.productCode,
    productName: item.productName,
    type: txnType,
    quantity: qty,
    stockAfter: nextQty,
    actor: authState.username || "unknown",
    note,
    createdAt: Date.now()
  });
  saveJSON(STORAGE.inventoryTransactions, inventoryTransactions);
  logActivity(
    "Kho",
    txnType === "in" ? "Nhập kho" : "Xuất kho",
    `${item.productCode} | SL: ${qty}${note ? ` | ${note}` : ""}`
  );
  closeInventoryTxnModal();
  showToast("Đã ghi nhận giao dịch kho.");
  renderAll();
});

els.openCustomerModalBtn.addEventListener("click", openCustomerModal);
els.openUserModalBtn.addEventListener("click", openUserModal);
els.closeUserModalBtn.addEventListener("click", closeUserModal);
els.userModalBackdrop.addEventListener("click", closeUserModal);
els.hrQuickSearch.addEventListener("input", renderUserTable);
els.hrFilterRole.addEventListener("change", renderUserTable);
els.hrFilterDept.addEventListener("change", renderUserTable);
els.hrFilterStatus.addEventListener("change", renderUserTable);
els.closeHrProfileBtn.addEventListener("click", closeHrProfileModal);
els.hrProfileModalBackdrop.addEventListener("click", closeHrProfileModal);

if (els.openPermissionModalBtn) {
  els.openPermissionModalBtn.addEventListener("click", () => {
    if (!can("canManageUsers")) return;
    openPermissionModal();
  });
}

if (els.closePermissionModalBtn) {
  els.closePermissionModalBtn.addEventListener("click", closePermissionModal);
}

if (els.permissionModalBackdrop) {
  els.permissionModalBackdrop.addEventListener("click", closePermissionModal);
}

if (els.permissionRoleSelect) {
  els.permissionRoleSelect.addEventListener("change", () => {
    permissionEditingRole = els.permissionRoleSelect.value || "staff";
    renderPermissionModal();
  });
}

if (els.permissionModal) {
  els.permissionModal.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;
    if (permissionEditingRole === "admin" || permissionEditingRole === "ceo") return;
    const rolePerms = rolePermissionsState[permissionEditingRole];
    if (!rolePerms) return;

    const featureKey = target.dataset.permFeature;
    if (featureKey) {
      rolePerms[featureKey] = target.checked;
      return;
    }

    const pageKey = target.dataset.permPage;
    if (pageKey) {
      const pages = new Set(rolePerms.pageAccess || []);
      if (target.checked) pages.add(pageKey);
      else pages.delete(pageKey);
      rolePerms.pageAccess = Array.from(pages);
    }
  });
}

if (els.savePermissionBtn) {
  els.savePermissionBtn.addEventListener("click", () => {
    rolePermissionsState = normalizeRolePermissions(rolePermissionsState);
    saveRolePermissionsState();
    renderPermissionModal();
    renderAll();
    showToast("Đã lưu cấu hình phân quyền.");
    logActivity("Nhân sự", "Cập nhật phân quyền", `Vai trò: ${permissionEditingRole}`);
    renderActivityTable();
  });
}

if (els.resetPermissionBtn) {
  els.resetPermissionBtn.addEventListener("click", () => {
    rolePermissionsState = normalizeRolePermissions(ROLES);
    saveRolePermissionsState();
    renderPermissionModal();
    renderAll();
    showToast("Đã đặt lại phân quyền mặc định.");
    logActivity("Nhân sự", "Đặt lại phân quyền", "Khôi phục mặc định theo hệ thống");
    renderActivityTable();
  });
}

els.hrFileUpload.addEventListener("change", () => {
  const files = Array.from(els.hrFileUpload.files);
  if (!files.length || !viewingHrUserId) return;
  const reads = files.map((file) => new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve({
      id: `hrf-${Date.now()}-${Math.random().toString(36).slice(2)}`,
      name: file.name,
      fileType: file.type || "application/octet-stream",
      data: e.target.result,
      uploadedAt: Date.now()
    });
    reader.readAsDataURL(file);
  }));
  Promise.all(reads).then((newFiles) => {
    if (!hrFiles[viewingHrUserId]) hrFiles[viewingHrUserId] = [];
    hrFiles[viewingHrUserId].push(...newFiles);
    saveJSON(STORAGE.hrFiles, hrFiles);
    renderHrFileList(viewingHrUserId);
    logActivity("Nhân sự", "Tải lên hồ sơ", `${newFiles.map((f) => f.name).join(", ")}`);
    els.hrProfileStatus.textContent = `Đã tải lên ${newFiles.length} tệp.`;
    els.hrFileUpload.value = "";
  });
});

els.hrProfileFileList.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  const fileId = target.dataset.fileId;
  const userId = target.dataset.userId;
  if (!fileId || !userId) return;

  if (target.classList.contains("hr-file-download-btn")) {
    const file = (hrFiles[userId] || []).find((f) => f.id === fileId);
    if (!file) return;
    const a = document.createElement("a");
    a.href = file.data;
    a.download = file.name;
    a.click();
    return;
  }

  if (target.classList.contains("hr-file-delete-btn")) {
    hrFiles[userId] = (hrFiles[userId] || []).filter((f) => f.id !== fileId);
    saveJSON(STORAGE.hrFiles, hrFiles);
    renderHrFileList(userId);
    els.hrProfileStatus.textContent = "Đã xóa tệp.";
  }
});

els.userBody.addEventListener("click", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  if (target.classList.contains("user-name-link")) {
    openHrProfileModal(target.dataset.userId);
    return;
  }
});

document.addEventListener("click", (e) => {
  if (!e.target.classList.contains("user-action-toggle")) {
    hideAllActionMenus();
  }
});

els.openCustomerModalBtn.addEventListener("click", openCustomerModal);
els.closeCustomerModalBtn.addEventListener("click", closeCustomerModal);
els.customerModalBackdrop.addEventListener("click", closeCustomerModal);
els.customerAddrProvince.addEventListener("change", () => {
  renderAddrDistricts(els.customerAddrProvince.value);
  composeAddressFromSelects();
});
els.customerAddrDistrict.addEventListener("change", () => {
  renderAddrWards(els.customerAddrProvince.value, els.customerAddrDistrict.value);
  composeAddressFromSelects();
});
els.customerAddrWard.addEventListener("change", composeAddressFromSelects);
els.applyCustomerFilterBtn.addEventListener("click", () => {
  applyCustomerFiltersFromInputs();
  renderCustomerTable();
});
els.customerDateRangeTrigger.addEventListener("click", () => {
  const isOpen = !els.customerDateRangePopover.classList.contains("hidden");
  if (isOpen) {
    closeCustomerDateRangePopover();
    return;
  }
  openCustomerDateRangePopover();
});
els.applyCustomerDateRangeBtn.addEventListener("click", () => {
  applyCustomerDateRangeFromInputs();
  closeCustomerDateRangePopover();
  renderCustomerTable();
});
els.customerDateRangePresets.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  const preset = target.dataset.rangePreset;
  if (!preset) return;
  applyCustomerDatePreset(preset);
  closeCustomerDateRangePopover();
  renderCustomerTable();
});
els.clearCustomerDateRangeBtn.addEventListener("click", () => {
  customerFilterState = {
    ...customerFilterState,
    start: "",
    end: ""
  };
  saveCustomerFilterState();
  els.customerFilterStartDate.value = "";
  els.customerFilterEndDate.value = "";
  closeCustomerDateRangePopover();
  renderCustomerTable();
});
els.resetCustomerFilterBtn.addEventListener("click", () => {
  resetCustomerFilters();
  renderCustomerTable();
});
els.customerQuickSearch.addEventListener("input", () => {
  setCustomerKeywordFilter(els.customerQuickSearch.value);
  renderCustomerTable();
});
els.exportCustomerExcelBtn.addEventListener("click", () => {
  exportFilteredCustomersToExcel();
});
els.exportCustomerPdfBtn.addEventListener("click", async () => {
  await exportFilteredCustomersToPdf();
});

document.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof Node)) return;
  if (els.customerDateRangeField.contains(target)) return;
  closeCustomerDateRangePopover();
});

if (els.activityBody) {
  els.activityBody.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (!can("canManageUsers")) return;

    const restoreRef = target.dataset.restoreRef;
    const restoreActivityId = target.dataset.restoreActivityId;
    if (!restoreRef && !restoreActivityId) return;

    let restored = false;
    if (restoreRef) {
      restored = restoreDeletedRecord(restoreRef);
    } else if (restoreActivityId) {
      restored = restoreActivityAction(restoreActivityId);
    }

    if (!restored) return;
    renderAll();
  });
}

if (els.applyActivityFilterBtn) {
  els.applyActivityFilterBtn.addEventListener("click", () => {
    let start = els.activityFilterStartDate?.value || "";
    let end = els.activityFilterEndDate?.value || "";
    if (start && end && start > end) {
      const temp = start;
      start = end;
      end = temp;
    }
    activityViewState.start = start;
    activityViewState.end = end;
    activityViewState.page = 1;
    renderActivityTable();
  });
}

if (els.resetActivityFilterBtn) {
  els.resetActivityFilterBtn.addEventListener("click", () => {
    activityViewState.start = today;
    activityViewState.end = today;
    activityViewState.page = 1;
    renderActivityTable();
  });
}

if (els.activityPrevPageBtn) {
  els.activityPrevPageBtn.addEventListener("click", () => {
    activityViewState.page = Math.max(1, activityViewState.page - 1);
    renderActivityTable();
  });
}

if (els.activityNextPageBtn) {
  els.activityNextPageBtn.addEventListener("click", () => {
    activityViewState.page += 1;
    renderActivityTable();
  });
}

els.backBtn.addEventListener("click", () => {
  if (pageHistory.length === 0) return;
  const prevPage = pageHistory.pop();
  setActivePage(prevPage || "home", { fromHistory: true });
  renderAll();
});

els.appMenu.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  const menuBtn = target.closest(".menu-item");
  if (!(menuBtn instanceof HTMLElement)) return;
  const pageKey = menuBtn.dataset.page;
  if (!pageKey) return;
  setActivePage(pageKey);
  renderAll();
});

window.addEventListener("popstate", () => {
  if (!authState.loggedIn) return;
  const pageFromLocation = getPageKeyFromLocation() || "home";
  setActivePage(pageFromLocation, { fromHistory: true, syncUrl: false });
  renderAll();
});

if (els.workflowSection) {
  els.workflowSection.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const detailBtn = target.closest("[data-workflow-target]");
    if (!(detailBtn instanceof HTMLElement)) return;
    const departmentKey = detailBtn.dataset.workflowTarget;
    if (!departmentKey) return;
    openWorkflowDepartmentDetail(departmentKey);
  });
}

if (els.workflowBackBtn) {
  els.workflowBackBtn.addEventListener("click", () => {
    closeWorkflowDepartmentDetail();
  });
}

if (els.policySection) {
  els.policySection.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const parentBtn = target.closest("[data-policy-parent]");
    if (parentBtn instanceof HTMLElement) {
      const parentKey = parentBtn.dataset.policyParent;
      if (!parentKey) return;
      togglePolicyParent(parentKey);
      return;
    }
    const childBtn = target.closest("[data-policy-child]");
    if (!(childBtn instanceof HTMLElement)) return;
    const childId = childBtn.dataset.policyChild;
    if (!childId) return;
    togglePolicyChild(childId);
  });
}

if (els.reportsSection) {
  els.reportsSection.addEventListener("click", async (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const consultantSortHeader = target.closest("[data-consultant-sort]");
    if (consultantSortHeader instanceof HTMLElement && activeReportDepartment === "consultant") {
      const sortKey = consultantSortHeader.dataset.consultantSort;
      if (!sortKey) return;
      toggleConsultantReportSort(sortKey);
      renderReportsPage();
      return;
    }
    const telesaleSortHeader = target.closest("[data-telesale-sort]");
    if (telesaleSortHeader instanceof HTMLElement && activeReportDepartment === "telesale") {
      const sortKey = telesaleSortHeader.dataset.telesaleSort;
      if (!sortKey) return;
      toggleTelesaleReportSort(sortKey);
      renderReportsPage();
      return;
    }
    const sortHeader = target.closest("[data-nurse-sort]");
    if (sortHeader instanceof HTMLElement && activeReportDepartment === "nurse") {
      const sortKey = sortHeader.dataset.nurseSort;
      if (!sortKey) return;
      toggleNurseReportSort(sortKey);
      renderNurseReportMatrix();
      return;
    }
    const nurseNameCell = target.closest("[data-nurse-name]");
    if (nurseNameCell instanceof HTMLElement && activeReportDepartment === "nurse") {
      const nurseName = decodeURIComponent(nurseNameCell.dataset.nurseName || "");
      if (!nurseName) return;
      showNurseDailyDetailModal(nurseName, reportFilterState.start, reportFilterState.end);
      return;
    }
    const folderBtn = target.closest("[data-report-dept]");
    if (!(folderBtn instanceof HTMLElement)) return;
    const departmentKey = folderBtn.dataset.reportDept;
    if (!departmentKey) return;
    await openReportDepartment(departmentKey);
  });

}

if (els.accountingSection) {
  els.accountingSection.addEventListener("click", async (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const folderBtn = target.closest("[data-accounting-folder]");
    if (folderBtn instanceof HTMLElement) {
      const folderKey = folderBtn.dataset.accountingFolder;
      if (!folderKey) return;
      openAccountingFolder(folderKey);
      return;
    }

    if (target.id === "openCashflowModalBtn") {
      if (!can("canSubmitReport")) {
        showToast("Bạn không có quyền tạo phiếu thu chi.", "warning");
        return;
      }
      els.accountingSection.querySelector("#cashflowModal")?.classList.remove("hidden");
      return;
    }

    if (target.id === "closeCashflowModalBtn" || target.id === "cashflowModalBackdrop") {
      els.accountingSection.querySelector("#cashflowModal")?.classList.add("hidden");
      return;
    }

    if (target.id === "openServicePayrollInfoBtn") {
      els.accountingSection.querySelector("#servicePayrollInfoModal")?.classList.remove("hidden");
      return;
    }

    if (target.id === "closeServicePayrollInfoBtn" || target.id === "servicePayrollInfoBackdrop") {
      els.accountingSection.querySelector("#servicePayrollInfoModal")?.classList.add("hidden");
      return;
    }

    if (target.id === "applyCashflowFilterBtn") {
      const startInput = els.accountingSection.querySelector("#cashflowFilterStartDate");
      const endInput = els.accountingSection.querySelector("#cashflowFilterEndDate");
      accountingCashflowFilterState = normalizeAccountingCashflowFilterState({
        start: startInput instanceof HTMLInputElement ? startInput.value : "",
        end: endInput instanceof HTMLInputElement ? endInput.value : ""
      });
      saveAccountingCashflowFilterState();
      renderAccountingPage();
      return;
    }

    if (target.id === "resetCashflowFilterBtn") {
      accountingCashflowFilterState = normalizeAccountingCashflowFilterState({
        start: `${today.slice(0, 7)}-01`,
        end: today
      });
      saveAccountingCashflowFilterState();
      renderAccountingPage();
      return;
    }

    if (target.id === "applyAttendanceFilterBtn") {
      const startInput = els.accountingSection.querySelector("#attendanceFilterStartDate");
      const endInput = els.accountingSection.querySelector("#attendanceFilterEndDate");
      accountingAttendanceFilterState = normalizeAccountingAttendanceFilterState({
        start: startInput instanceof HTMLInputElement ? startInput.value : "",
        end: endInput instanceof HTMLInputElement ? endInput.value : ""
      });
      saveAccountingAttendanceFilterState();
      renderAccountingPage();
      return;
    }

    if (target.id === "resetAttendanceFilterBtn") {
      accountingAttendanceFilterState = normalizeAccountingAttendanceFilterState({
        start: `${today.slice(0, 7)}-01`,
        end: today
      });
      saveAccountingAttendanceFilterState();
      renderAccountingPage();
      return;
    }

    if (target.id === "applyServicePayrollFilterBtn") {
      const startInput = els.accountingSection.querySelector("#servicePayrollFilterStartDate");
      const endInput = els.accountingSection.querySelector("#servicePayrollFilterEndDate");
      accountingServicePayrollFilterState = normalizeAccountingServicePayrollFilterState({
        start: startInput instanceof HTMLInputElement ? startInput.value : "",
        end: endInput instanceof HTMLInputElement ? endInput.value : ""
      });
      saveAccountingServicePayrollFilterState();
      renderAccountingPage();
      return;
    }

    if (target.id === "resetServicePayrollFilterBtn") {
      accountingServicePayrollFilterState = normalizeAccountingServicePayrollFilterState({
        start: `${today.slice(0, 7)}-01`,
        end: today
      });
      saveAccountingServicePayrollFilterState();
      renderAccountingPage();
      return;
    }

    if (target.id === "syncAttendanceBtn") {
      if (!can("canSyncData")) {
        showToast("Bạn không có quyền đồng bộ dữ liệu chấm công.", "warning");
        return;
      }
      const sourceTypeInput = els.accountingSection.querySelector("#attendanceSourceType");
      const vendorInput = els.accountingSection.querySelector("#attendanceVendorType");
      const sourceUrlInput = els.accountingSection.querySelector("#attendanceSourceUrl");
      accountingAttendanceSource.type = sourceTypeInput instanceof HTMLSelectElement ? sourceTypeInput.value : "sheet";
      accountingAttendanceSource.vendor = vendorInput instanceof HTMLSelectElement ? vendorInput.value : "generic";
      accountingAttendanceSource.url = sourceUrlInput instanceof HTMLInputElement ? sourceUrlInput.value.trim() : "";
      saveAccountingAttendanceSource();
      configureAttendanceAutoSync();
      await runAttendanceSync();
      return;
    }

    if (target.id === "clearAttendanceBtn") {
      if (!can("canSyncData")) {
        showToast("Bạn không có quyền xóa dữ liệu chấm công.", "warning");
        return;
      }
      accountingAttendanceEntries = [];
      saveJSON(STORAGE.accountingAttendance, accountingAttendanceEntries);
      accountingAttendanceSource.lastWarning = "";
      saveAccountingAttendanceSource();
      renderAccountingPage();
      showToast("Đã xóa dữ liệu chấm công.", "warning");
      logActivity("Kế toán", "Xóa dữ liệu chấm công", "Xóa toàn bộ bản ghi tính công");
      renderActivityTable();
      return;
    }

    if (target.id === "saveCashflowBtn") {
      if (!can("canSubmitReport")) {
        showToast("Bạn không có quyền tạo phiếu thu chi.", "warning");
        return;
      }
      const typeInput = els.accountingSection.querySelector("#cashflowType");
      const dateInput = els.accountingSection.querySelector("#cashflowDate");
      const categoryInput = els.accountingSection.querySelector("#cashflowCategory");
      const counterpartyInput = els.accountingSection.querySelector("#cashflowCounterparty");
      const amountInput = els.accountingSection.querySelector("#cashflowAmount");
      const methodInput = els.accountingSection.querySelector("#cashflowMethod");
      const statusInput = els.accountingSection.querySelector("#cashflowStatus");
      const creatorInput = els.accountingSection.querySelector("#cashflowCreator");
      const contentInput = els.accountingSection.querySelector("#cashflowContent");
      const statusLabel = els.accountingSection.querySelector("#cashflowModalStatus");

      const type = typeInput instanceof HTMLSelectElement ? typeInput.value : "income";
      const date = dateInput instanceof HTMLInputElement ? dateInput.value : today;
      const category = categoryInput instanceof HTMLInputElement ? categoryInput.value.trim() : "";
      const counterparty = counterpartyInput instanceof HTMLInputElement ? counterpartyInput.value.trim() : "";
      const amount = amountInput instanceof HTMLInputElement ? Number(amountInput.value || 0) : 0;
      const method = methodInput instanceof HTMLSelectElement ? methodInput.value : "Tiền mặt";
      const status = statusInput instanceof HTMLSelectElement ? statusInput.value : "pending";
      const creator = creatorInput instanceof HTMLInputElement ? creatorInput.value.trim() : "";
      const content = contentInput instanceof HTMLTextAreaElement ? contentInput.value.trim() : "";

      if (!date || !category || !counterparty || !content || amount <= 0) {
        if (statusLabel instanceof HTMLElement) statusLabel.textContent = "Vui lòng nhập đầy đủ ngày, khoản mục, đối tượng, nội dung và số tiền hợp lệ.";
        return;
      }

      accountingCashflowEntries.unshift({
        id: `cf-${Date.now()}`,
        date,
        type: type === "expense" ? "expense" : "income",
        voucherCode: generateCashflowVoucherCode(type, date),
        category,
        counterparty,
        content,
        amount,
        method,
        creator: creator || getCurrentUser()?.fullName || getCurrentUser()?.username || "Hệ thống",
        status,
        createdAt: Date.now()
      });
      saveJSON(STORAGE.accountingCashflow, accountingCashflowEntries);
      els.accountingSection.querySelector("#cashflowModal")?.classList.add("hidden");
      renderAccountingPage();
      showToast(type === "expense" ? "Đã tạo phiếu chi." : "Đã tạo phiếu thu.");
      logActivity("Kế toán", type === "expense" ? "Tạo phiếu chi" : "Tạo phiếu thu", `${category} | ${amount.toLocaleString("vi-VN")} đ`);
      renderActivityTable();
    }
  });
}

if (els.accountingSection) {
  els.accountingSection.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (target.id === "attendanceSourceType" && target instanceof HTMLSelectElement) {
      accountingAttendanceSource.type = target.value;
      saveAccountingAttendanceSource();
      configureAttendanceAutoSync();
      return;
    }
    if (target.id === "attendanceVendorType" && target instanceof HTMLSelectElement) {
      accountingAttendanceSource.vendor = target.value;
      saveAccountingAttendanceSource();
      return;
    }
    if (target.id === "attendanceSourceUrl" && target instanceof HTMLInputElement) {
      accountingAttendanceSource.url = target.value.trim();
      saveAccountingAttendanceSource();
      configureAttendanceAutoSync();
      return;
    }
    if (target.id === "attendanceAutoSyncEnabled" && target instanceof HTMLInputElement) {
      accountingAttendanceSource.autoSyncEnabled = target.checked;
      saveAccountingAttendanceSource();
      configureAttendanceAutoSync();
      return;
    }
    if (target.id === "attendanceAutoSyncMinutes" && target instanceof HTMLSelectElement) {
      accountingAttendanceSource.autoSyncMinutes = Number(target.value || 10);
      saveAccountingAttendanceSource();
      configureAttendanceAutoSync();
    }
  });
}

if (els.newsSection) {
  els.newsSection.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;

    const detailOpenTarget = target.closest("[data-news-open-detail-type]");
    if (detailOpenTarget instanceof HTMLElement) {
      const detailType = detailOpenTarget.dataset.newsOpenDetailType || "post";
      const detailId = detailOpenTarget.dataset.newsOpenDetailId || "";
      const detailImageIndex = Number(detailOpenTarget.dataset.newsDetailImageIndex || "0");
      if (detailId) {
        openNewsContentDetail(detailType, detailId, detailImageIndex);
        return;
      }
    }

    const removeAttachmentBtn = target.closest("[data-news-remove-attachment]");
    if (removeAttachmentBtn instanceof HTMLElement) {
      const context = removeAttachmentBtn.dataset.newsRemoveAttachment;
      const index = Number(removeAttachmentBtn.dataset.newsAttachmentIndex);
      if (Number.isNaN(index) || index < 0) return;
      if (context === "event") {
        pendingNewsEventAttachments = pendingNewsEventAttachments.filter((_, idx) => idx !== index);
        renderEditorAttachmentList(els.newsEventAttachmentList, pendingNewsEventAttachments, "event");
      } else {
        pendingNewsPostAttachments = pendingNewsPostAttachments.filter((_, idx) => idx !== index);
        renderEditorAttachmentList(els.newsComposerAttachmentList, pendingNewsPostAttachments, "post");
      }
      return;
    }

    const postToggleBtn = target.closest("[data-news-post-toggle]");
    if (postToggleBtn instanceof HTMLElement) {
      const menu = document.querySelector(`[data-news-post-menu="${postToggleBtn.dataset.newsPostToggle}"]`);
      if (menu instanceof HTMLElement) openActionMenuAtToggle(postToggleBtn, menu);
      return;
    }

    const eventToggleBtn = target.closest("[data-news-event-toggle]");
    if (eventToggleBtn instanceof HTMLElement) {
      const menu = document.querySelector(`[data-news-event-menu="${eventToggleBtn.dataset.newsEventToggle}"]`);
      if (menu instanceof HTMLElement) openActionMenuAtToggle(eventToggleBtn, menu);
      return;
    }

    const editPostBtn = target.closest("[data-news-post-edit]");
    if (editPostBtn instanceof HTMLElement) {
      editNewsPost(editPostBtn.dataset.newsPostEdit || "");
      return;
    }

    const deletePostBtn = target.closest("[data-news-post-delete]");
    if (deletePostBtn instanceof HTMLElement) {
      deleteNewsPost(deletePostBtn.dataset.newsPostDelete || "");
      return;
    }

    const editEventBtn = target.closest("[data-news-event-edit]");
    if (editEventBtn instanceof HTMLElement) {
      editNewsEvent(editEventBtn.dataset.newsEventEdit || "");
      return;
    }

    const deleteEventBtn = target.closest("[data-news-event-delete]");
    if (deleteEventBtn instanceof HTMLElement) {
      deleteNewsEvent(deleteEventBtn.dataset.newsEventDelete || "");
      return;
    }

    const voteDetailBtn = target.closest("[data-news-vote-detail-event]");
    if (voteDetailBtn instanceof HTMLElement) {
      openNewsVoteDetail(voteDetailBtn.dataset.newsVoteDetailEvent || "", voteDetailBtn.dataset.newsVoteDetailChoice || "yes");
      return;
    }

    const voteTarget = target.closest("[data-news-vote-event]");
    const voteEventId = voteTarget instanceof HTMLElement ? voteTarget.dataset.newsVoteEvent : "";
    const voteChoice = voteTarget instanceof HTMLElement ? voteTarget.dataset.newsVoteChoice : "";
    if (voteEventId && voteChoice) {
      voteNewsEvent(voteEventId, voteChoice);
      return;
    }

    const actionTarget = target.closest("[data-news-action]");
    if (!(actionTarget instanceof HTMLElement)) return;
    const action = actionTarget.dataset.newsAction;
    if (!action) return;

    if (action === "open-composer") {
      openNewsComposer();
      return;
    }
    if (action === "submit-post") {
      submitNewsPost();
      return;
    }
    if (action === "create-event") {
      createNewsEvent();
      return;
    }
    if (action === "attach-department") {
      openNewsComposer();
      attachNewsDepartment();
    }
  });
}

if (els.newsComposerAttachmentInput) {
  els.newsComposerAttachmentInput.addEventListener("change", async () => {
    const files = Array.from(els.newsComposerAttachmentInput.files || []);
    if (!files.length) return;
    const attachments = await readFilesAsAttachments(files);
    pendingNewsPostAttachments = [...pendingNewsPostAttachments, ...attachments].slice(0, 12);
    renderEditorAttachmentList(els.newsComposerAttachmentList, pendingNewsPostAttachments, "post");
    els.newsComposerAttachmentInput.value = "";
    showToast(`Đã thêm ${attachments.length} tệp đính kèm cho bài tin.`, "info");
  });
}

if (els.newsEventAttachmentInput) {
  els.newsEventAttachmentInput.addEventListener("change", async () => {
    const files = Array.from(els.newsEventAttachmentInput.files || []);
    if (!files.length) return;
    const attachments = await readFilesAsAttachments(files);
    pendingNewsEventAttachments = [...pendingNewsEventAttachments, ...attachments].slice(0, 12);
    renderEditorAttachmentList(els.newsEventAttachmentList, pendingNewsEventAttachments, "event");
    els.newsEventAttachmentInput.value = "";
    showToast(`Đã thêm ${attachments.length} tệp đính kèm cho sự kiện.`, "info");
  });
}

if (els.saveNewsEventBtn) {
  els.saveNewsEventBtn.addEventListener("click", () => {
    saveNewsEvent();
  });
}

if (els.closeNewsEventModalBtn) {
  els.closeNewsEventModalBtn.addEventListener("click", closeNewsEventModal);
}

if (els.newsEventModalBackdrop) {
  els.newsEventModalBackdrop.addEventListener("click", closeNewsEventModal);
}

if (els.closeNewsVoteDetailBtn) {
  els.closeNewsVoteDetailBtn.addEventListener("click", closeNewsVoteDetail);
}

if (els.newsVoteDetailModalBackdrop) {
  els.newsVoteDetailModalBackdrop.addEventListener("click", closeNewsVoteDetail);
}

if (els.closeNewsContentDetailBtn) {
  els.closeNewsContentDetailBtn.addEventListener("click", closeNewsContentDetail);
}

if (els.newsContentDetailBackdrop) {
  els.newsContentDetailBackdrop.addEventListener("click", closeNewsContentDetail);
}

if (els.accountingBackBtn) {
  els.accountingBackBtn.addEventListener("click", () => {
    closeAccountingFolder();
  });
}

if (els.reportsBackBtn) {
  els.reportsBackBtn.addEventListener("click", () => {
    closeReportDepartment();
  });
}

if (els.applyReportsFilterBtn) {
  els.applyReportsFilterBtn.addEventListener("click", async () => {
    let start = els.reportsStartDate.value || "";
    let end = els.reportsEndDate.value || "";
    if (start && end && start > end) {
      const temp = start;
      start = end;
      end = temp;
    }
    reportFilterState.start = start;
    reportFilterState.end = end;
    if (activeReportDepartment === "nurse") {
      await syncTelegramBeforeNurseReportRender();
    }
    renderReportsPage();
  });
}

if (els.resetReportsFilterBtn) {
  els.resetReportsFilterBtn.addEventListener("click", async () => {
    reportFilterState = { start: `${today.slice(0, 7)}-01`, end: today };
    marketingReportState.marketing = "";
    consultantReportState.consultant = "";
    consultantReportState.sortKey = "date";
    consultantReportState.direction = "desc";
    telesaleReportState.sale = "";
    telesaleReportState.sortKey = "date";
    telesaleReportState.direction = "desc";
    if (activeReportDepartment === "nurse") {
      await syncTelegramBeforeNurseReportRender();
    }
    renderReportsPage();
  });
}

if (els.reportsConsultantFilter) {
  els.reportsConsultantFilter.addEventListener("change", () => {
    consultantReportState.consultant = els.reportsConsultantFilter.value || "";
    renderReportsPage();
  });
}

if (els.reportsMarketingFilter) {
  els.reportsMarketingFilter.addEventListener("change", () => {
    marketingReportState.marketing = els.reportsMarketingFilter.value || "";
    renderReportsPage();
  });
}

if (els.reportsTelesaleFilter) {
  els.reportsTelesaleFilter.addEventListener("change", () => {
    telesaleReportState.sale = els.reportsTelesaleFilter.value || "";
    renderReportsPage();
  });
}

if (els.reportsSyncMetricsBtn) {
  els.reportsSyncMetricsBtn.addEventListener("click", () => {
    if (!activeReportDepartment) return;
    metricsFilterState.start = reportFilterState.start;
    metricsFilterState.end = reportFilterState.end;
    metricsFilterState.department = activeReportDepartment;
    setActivePage("metrics");
    renderAll();
    showToast("Đã đồng bộ bộ lọc báo cáo sang Trang chỉ số.");
  });
}

if (els.exportReportsExcelBtn) {
  els.exportReportsExcelBtn.addEventListener("click", () => {
    exportReportDetailExcel();
  });
}

if (els.exportReportsPdfBtn) {
  els.exportReportsPdfBtn.addEventListener("click", async () => {
    await exportReportDetailPdf();
  });
}

els.mainContent.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  const goPage = target.dataset.goPage;
  if (!goPage) return;
  setActivePage(goPage);
  renderAll();
});

els.timePreset.addEventListener("change", () => {
  const preset = els.timePreset.value;
  if (preset === "custom") {
    filterState = { ...filterState, preset: "custom" };
    logActivity("Bộ lọc", "Chọn chế độ tùy chỉnh", "Đã chuyển sang lọc theo khoảng ngày tùy chỉnh");
    return;
  }
  applyPresetFilter(preset);
  logActivity("Bộ lọc", "Áp preset thời gian", `Preset: ${preset} | Từ ${filterState.start} đến ${filterState.end}`);
  renderAll();
});

els.applyFilterBtn.addEventListener("click", () => {
  applyCustomRange();
  logActivity("Bộ lọc", "Áp dụng khoảng ngày", `Từ ${filterState.start} đến ${filterState.end}`);
  renderAll();
});

els.resetFilterBtn.addEventListener("click", () => {
  applyPresetFilter("today");
  logActivity("Bộ lọc", "Đặt lại bộ lọc", "Đặt bộ lọc về hôm nay");
  renderAll();
});

els.applyMetricsFilterBtn.addEventListener("click", () => {
  let start = els.metricsStartDate.value || "";
  let end = els.metricsEndDate.value || "";
  if (start && end && start > end) {
    const temp = start;
    start = end;
    end = temp;
  }
  metricsFilterState.start = start;
  metricsFilterState.end = end;
  metricsFilterState.department = els.metricsDepartmentFilter.value || "all";
  renderDepartmentMetrics();
});

els.resetMetricsFilterBtn.addEventListener("click", () => {
  metricsFilterState = { start: filterState.start, end: filterState.end, department: "all" };
  renderMetricsFilterControls();
  renderDepartmentMetrics();
});

els.metricsImportFileBtn.addEventListener("click", async () => {
  const file = els.metricsDataFile.files?.[0];
  if (!file) {
    showToast("Vui lòng chọn file dữ liệu trước khi tải.", "warning");
    return;
  }
  const text = await file.text();
  let importedRows = [];
  try {
    if (file.name.toLowerCase().endsWith(".json")) {
      const parsed = JSON.parse(text);
      importedRows = Array.isArray(parsed) ? parsed : [];
    } else {
      importedRows = parseCsvText(text);
    }
  } catch (err) {
    showToast(`File không hợp lệ: ${err.message}`, "error");
    return;
  }
  const added = mergeImportedSchedules(importedRows);
  if (!added) {
    showToast("Không có bản ghi hợp lệ để nhập.", "warning");
    return;
  }
  renderAll();
  showToast(`Đã nhập ${added} bản ghi lịch từ file.`);
  logActivity("Chỉ số", "Nhập dữ liệu file", `Số bản ghi: ${added}`);
});

els.metricsSyncSheetBtn.addEventListener("click", async () => {
  const rawUrl = (els.metricsSheetUrl.value || "").trim();
  if (!rawUrl) {
    showToast("Vui lòng nhập link Google Sheets/CSV/JSON.", "warning");
    return;
  }
  const url = normalizeSheetUrl(rawUrl);
  try {
    showToast("Đang đồng bộ dữ liệu từ link...", "info");
    const response = await fetch(url, { method: "GET" });
    if (!response.ok) throw new Error(`Không tải được dữ liệu (${response.status})`);
    const text = await response.text();
    let importedRows = [];
    if (text.trim().startsWith("[") || text.trim().startsWith("{")) {
      const parsed = JSON.parse(text);
      importedRows = Array.isArray(parsed) ? parsed : [];
    } else {
      importedRows = parseCsvText(text);
    }
    const added = mergeImportedSchedules(importedRows);
    if (!added) {
      showToast("Không có bản ghi hợp lệ từ link.", "warning");
      return;
    }
    renderAll();
    showToast(`Đồng bộ thành công ${added} bản ghi từ link.`);
    logActivity("Chỉ số", "Đồng bộ link", `Endpoint: ${rawUrl} | Bản ghi: ${added}`);
  } catch (err) {
    showToast(`Lỗi đồng bộ link: ${err.message}`, "error");
  }
});

// ── Schedule event listeners ─────────────────────────────────────────────────
els.openScheduleModalBtn.addEventListener("click", () => openScheduleModal(null));
els.closeScheduleModalBtn.addEventListener("click", closeScheduleModal);
els.scheduleModalBackdrop.addEventListener("click", closeScheduleModal);

els.exportScheduleExcelBtn.addEventListener("click", exportScheduleExcel);
els.exportSchedulePdfBtn.addEventListener("click", exportSchedulePdf);

els.applyScheduleFilterBtn.addEventListener("click", () => {
  scheduleFilterState.month = els.scheduleFilterMonth.value || "";
  scheduleFilterState.status = els.scheduleFilterStatus.value || "";
  scheduleFilterState.staff = els.scheduleFilterStaff.value || "all";
  scheduleFilterState.source = els.scheduleFilterSource.value.trim();
  scheduleFilterState.keyword = els.scheduleSearch.value.trim();
  renderScheduleTable();
});

els.resetScheduleFilterBtn.addEventListener("click", () => {
  scheduleFilterState = { month: today.slice(0, 7), status: "", staff: "all", source: "", keyword: "" };
  renderScheduleStaffControls();
  els.scheduleFilterMonth.value = today.slice(0, 7);
  els.scheduleFilterStatus.value = "";
  els.scheduleFilterStaff.value = "all";
  els.scheduleFilterSource.value = "";
  els.scheduleSearch.value = "";
  renderScheduleTable();
});

els.scheduleSearch.addEventListener("input", () => {
  scheduleFilterState.keyword = els.scheduleSearch.value.trim();
  renderScheduleTable();
});

els.scheduleBody.addEventListener("dblclick", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  const cell = target.closest("td.schedule-editable");
  if (!(cell instanceof HTMLTableCellElement)) return;
  const row = cell.closest("tr[data-sch-id]");
  if (!(row instanceof HTMLTableRowElement)) return;
  const id = row.dataset.schId;
  const field = cell.dataset.field;
  if (!id || !field) return;
  const record = schedules.find((item) => item.id === id);
  if (!record) return;
  openScheduleInlineEditor(cell, record, field);
});

els.scheduleBody.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  const row = target.closest("tr[data-sch-id]");
  const id = row instanceof HTMLTableRowElement ? row.dataset.schId : undefined;
  if (!id) return;
  if (target.classList.contains("schedule-edit-btn")) {
    openScheduleModal(id);
  } else if (target.classList.contains("schedule-delete-btn") || event.altKey) {
    const s = schedules.find((x) => x.id === id);
    if (!s) return;
    if (!confirm(`Xóa lịch của "${s.customerName}" ngày ${s.registrationDate}?`)) return;
    schedules = schedules.filter((x) => x.id !== id);
    saveJSON(STORAGE.schedule, schedules);
    renderScheduleTable();
    renderCustomerCarePage();
    showToast("Đã xóa lịch.");
    logActivity("Lịch KH", "Xóa lịch", `${s.customerName} | ${s.registrationDate}`, { restoreAction: { kind: "schedule-delete", deletedSchedule: clonePlain(s) } });
  }
});

els.saveScheduleBtn.addEventListener("click", () => {
  const customerName = els.scheduleName.value.trim();
  if (!customerName) { showToast("Vui lòng nhập tên khách hàng.", "warning"); return; }
  const entry = {
    id: editingScheduleId || `sc-${Date.now()}`,
    registrationDate: els.scheduleRegDate.value || today,
    appointmentTime: els.scheduleTime.value.trim(),
    customerName,
    phone: els.schedulePhone.value.trim(),
    address: els.scheduleAddress.value.trim(),
    motherAge: els.scheduleMotherAge.value.trim(),
    birthHistory: els.scheduleBirthHistory.value.trim(),
    babyBirthday: els.scheduleBabyBirthday.value.trim(),
    priority: "",
    service: els.scheduleService.value.trim(),
    stage: els.scheduleStage.value.trim(),
    motherCondition: els.scheduleMotherCondition.value.trim(),
    babyCondition: els.scheduleBabyCondition.value.trim(),
    consultant: els.scheduleConsultant.value.trim(),
    nurse: els.scheduleNurse.value.trim(),
    saleStaff: els.scheduleSale.value.trim(),
    experiencePrice: Number(els.scheduleExpPrice.value) || 0,
    sessionDuration: els.scheduleSessionDuration.value.trim(),
    source: els.scheduleSource.value.trim(),
    contractAmount: Number(els.scheduleContractAmount.value) || 0,
    status: els.scheduleStatus.value || "pending",
    note: els.scheduleNote.value.trim(),
    updatedAt: Date.now(),
    createdAt: editingScheduleId ? (schedules.find((x) => x.id === editingScheduleId)?.createdAt || Date.now()) : Date.now()
  };
  if (editingScheduleId) {
    const previousSchedule = clonePlain(schedules.find((x) => x.id === editingScheduleId) || null);
    const idx = schedules.findIndex((x) => x.id === editingScheduleId);
    if (idx !== -1) schedules[idx] = entry;
    logActivity("Lịch KH", "Cập nhật lịch", `${customerName} | ${entry.registrationDate}`, previousSchedule ? { restoreAction: { kind: "schedule-edit", previousSchedule } } : {});
  } else {
    schedules.unshift(entry);
    logActivity("Lịch KH", "Thêm lịch mới", `${customerName} | ${entry.registrationDate}`);
  }
  saveJSON(STORAGE.schedule, schedules);
  closeScheduleModal();
  renderScheduleTable();
  renderCustomerCarePage();
  showToast(editingScheduleId ? "Đã cập nhật lịch." : "Đã thêm lịch mới.");
});

if (els.applyCareFilterBtn) {
  els.applyCareFilterBtn.addEventListener("click", () => {
    let start = els.careFilterStartDate.value || "";
    let end = els.careFilterEndDate.value || "";
    if (start && end && start > end) {
      const temp = start;
      start = end;
      end = temp;
    }
    customerCareFilterState = normalizeCustomerCareFilterState({
      ...customerCareFilterState,
      start,
      end,
      staff: els.careFilterStaff.value || "all",
      status: els.careFilterStatus.value || "all",
      source: els.careFilterSource.value || "all",
      progress: els.careFilterProgress.value || "all",
      keyword: (els.careSearch.value || "").trim()
    });
    saveCustomerCareFilterState();
    renderCustomerCareTable();
  });
}

if (els.resetCareFilterBtn) {
  els.resetCareFilterBtn.addEventListener("click", () => {
    customerCareFilterState = normalizeCustomerCareFilterState({
      start: "",
      end: "",
      staff: "all",
      status: "all",
      source: "all",
      progress: "all",
      keyword: ""
    });
    saveCustomerCareFilterState();
    renderCustomerCarePage();
  });
}

if (els.careSearch) {
  els.careSearch.addEventListener("input", () => {
    customerCareFilterState = normalizeCustomerCareFilterState({
      ...customerCareFilterState,
      keyword: els.careSearch.value.trim()
    });
    saveCustomerCareFilterState();
    renderCustomerCareTable();
  });
}

if (els.careBody) {
  els.careBody.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const saveBtn = target.closest(".care-save-btn");
    if (!(saveBtn instanceof HTMLElement)) return;
    const careKey = saveBtn.dataset.careKey;
    if (!careKey) return;
    const row = saveBtn.closest("tr");
    if (!(row instanceof HTMLTableRowElement)) return;
    saveCustomerCareProgressFromRow(careKey, row);
  });
}

if (els.exportCareExcelBtn) {
  els.exportCareExcelBtn.addEventListener("click", () => {
    exportFilteredCustomerCareToExcel();
  });
}

if (els.exportCarePdfBtn) {
  els.exportCarePdfBtn.addEventListener("click", async () => {
    await exportFilteredCustomerCareToPdf();
  });
}

// ── End Schedule event listeners ─────────────────────────────────────────────

els.logoUpload.addEventListener("change", () => {
  const file = els.logoUpload.files?.[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = () => {
    const result = typeof reader.result === "string" ? reader.result : "";
    if (!result) return;
    localStorage.setItem(STORAGE.logo, result);
    setBrandLogo(result);
    showToast("Đã cập nhật logo thương hiệu.");
    logActivity("Hệ thống", "Cập nhật logo", "Thay đổi logo thương hiệu trên header");
    renderActivityTable();
  };
  reader.readAsDataURL(file);
});

async function fetchRemoteReports(type, url) {
  if (!url) throw new Error("Thiếu endpoint URL");
  const response = await fetch(url, { method: "GET" });
  if (!response.ok) throw new Error(`Không thể GET từ ${type}`);

  if (type === "sheet") {
    // Handle CSV for Google Sheets
    const text = await response.text();
    const parsed = parseCsvText(text);
    // Normalize CSV data to report format
    return parsed.map(row => ({
      date: row.date || row.ngay || today,
      department: row.department || row.phongban || row.bophan || "Chưa xác định",
      completion: Number(row.completion || row.hoanthanh || row.hoan_thanh || 0),
      quality: Number(row.quality || row.chatluong || row.chat_luong || 0),
      issues: Number(row.issues || row.vandevande || row.vande || 0),
      submitter: row.submitter || row.nguoigui || row.nguoi_gui || "Unknown",
      updatedAt: row.updatedat ? new Date(row.updatedat).getTime() : Date.now()
    })).filter(report => report.completion > 0 || report.quality > 0);
  } else {
    // Handle JSON for API
    const data = await response.json();
    if (!Array.isArray(data)) throw new Error("Dữ liệu trả về phải là mảng report");
    return data;
  }
}

async function postRemoteReport(type, url, payload) {
  if (!url) throw new Error("Thiếu endpoint URL");
  if (type === "sheet") {
    // For Google Sheets, POST is not supported in this simple implementation
    // Would require Google Apps Script web app
    throw new Error("POST không hỗ trợ cho Google Sheets. Chỉ dùng để đồng bộ dữ liệu.");
  }
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  if (!response.ok) throw new Error(`Không thể POST lên ${type}`);
}

els.loginBtn.addEventListener("click", async () => {
  try {
    await syncUsersFromRemote(false);
  } catch (err) {
    showToast(`Không thể tải tài khoản cloud: ${err.message}`, "warning");
  }

  const username = els.loginUsername.value.trim().toLowerCase();
  const password = els.loginPassword.value;
  const user = users.find((u) => u.username.toLowerCase() === username);

  if (!user || user.password !== password) {
    els.authMessage.textContent = "Tên đăng nhập hoặc mật khẩu không hợp lệ";
    return;
  }

  if ((user.status || "active") === "suspended") {
    els.authMessage.textContent = "Tài khoản đang tạm dừng. Vui lòng liên hệ quản trị viên.";
    return;
  }

  authState = {
    loggedIn: true,
    role: getRolePermissions(user.roleKey).label,
    username: user.username,
    userId: user.id
  };
  saveLoginPrefs(Boolean(els.loginRemember?.checked), user.username, password);
  const requestedPage = getPageKeyFromLocation();
  activePage = requestedPage && canAccessPage(requestedPage) ? requestedPage : "news";
  pageHistory = [];
  saveJSON(STORAGE.auth, authState);
  logActivity("Xác thực", "Đăng nhập", `Người dùng: ${user.username}`);
  renderAll();
});

if (els.toggleLoginPasswordBtn) {
  els.toggleLoginPasswordBtn.addEventListener("click", () => {
    const toText = els.loginPassword.type === "password";
    els.loginPassword.type = toText ? "text" : "password";
    els.toggleLoginPasswordBtn.textContent = toText ? "🙈" : "👁";
    els.toggleLoginPasswordBtn.setAttribute("aria-label", toText ? "Ẩn mật khẩu" : "Hiện mật khẩu");
  });
}

if (els.loginRemember) {
  els.loginRemember.addEventListener("change", () => {
    if (els.loginRemember.checked) return;
    saveLoginPrefs(false, "", "");
  });
}

els.logoutBtn.addEventListener("click", () => {
  performLogout();
});

els.menuLogoutBtn.addEventListener("click", () => {
  performLogout();
});

els.saveCustomerBtn.addEventListener("click", () => {
  const name = els.customerName.value.trim();
  const contactPerson = els.customerContactPerson.value.trim();
  const phone = els.customerPhone.value.trim();
  const email = els.customerEmail.value.trim();
  const address = els.customerAddress.value.trim();
  const tier = els.customerTier.value;
  const status = els.customerStatus.value;
  const owner = els.customerOwner.value;
  const source = els.customerSource.value;
  const demand = els.customerDemand.value.trim();
  const note = els.customerNote.value.trim();

  if (!name || !phone || !contactPerson) {
    els.customerStatusMessage.textContent = "Vui lòng nhập tên khách hàng, người liên hệ và số điện thoại.";
    return;
  }

  if (editingCustomerId) {
    customers = customers.map((item) => {
      if (item.id !== editingCustomerId) return item;
      return {
        ...item,
        name,
        contactPerson,
        phone,
        email,
        address,
        tier,
        status,
        owner,
        source,
        demand,
        note,
        updatedAt: Date.now()
      };
    });
    saveJSON(STORAGE.customers, customers);
    logActivity("Khách hàng", "Cập nhật hồ sơ khách hàng", `${name} | Trạng thái: ${status}`);
    showToast("Đã cập nhật hồ sơ khách hàng.");
  } else {
    customers.unshift({
      id: `c-${Date.now()}`,
      name,
      contactPerson,
      phone,
      email,
      address,
      tier,
      status,
      owner,
      source,
      demand,
      note,
      updatedAt: Date.now()
    });
    saveJSON(STORAGE.customers, customers);
    logActivity("Khách hàng", "Thêm hồ sơ khách hàng", `${name} | Trạng thái: ${status}`);
    showToast("Đã lưu hồ sơ khách hàng.");
  }

  resetCustomerForm();
  closeCustomerModal();
  renderCustomerTable();
});

els.customerBody.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  const customerId = target.dataset.customerId;
  if (!customerId) return;

  if (target.classList.contains("customer-edit-btn")) {
    const customer = customers.find((item) => item.id === customerId);
    if (!customer) return;

    editingCustomerId = customer.id;
    els.customerModalTitle.textContent = `Chỉnh sửa: ${customer.name}`;
    els.saveCustomerBtn.textContent = "Cập nhật khách hàng";
    els.customerName.value = customer.name || "";
    els.customerContactPerson.value = customer.contactPerson || "";
    els.customerPhone.value = customer.phone || "";
    els.customerEmail.value = customer.email || "";
    els.customerAddress.value = customer.address || "";
    els.customerTier.value = customer.tier || "Tiêu chuẩn";
    els.customerStatus.value = customer.status || "Đã gọi";
    renderCustomerOwnerOptions(customer.owner || "Chưa gán");
    els.customerOwner.value = customer.owner || "Chưa gán";
    els.customerSource.value = customer.source || "Nhập tay";
    els.customerDemand.value = customer.demand || "";
    els.customerNote.value = customer.note || "";
    openCustomerModal();
    return;
  }

  if (!target.classList.contains("customer-delete-btn")) return;

  const deletedCustomer = customers.find((item) => item.id === customerId);
  customers = customers.filter((item) => item.id !== customerId);
  saveJSON(STORAGE.customers, customers);
  if (deletedCustomer) {
    const restoreRef = addToRecycleBin("customer", deletedCustomer, `Khách hàng: ${deletedCustomer.name}`);
    logActivity("Khách hàng", "Xóa hồ sơ khách hàng", deletedCustomer.name, { restoreRef });
  }
  els.customerStatusMessage.textContent = "Đã xóa khách hàng.";
  renderCustomerTable();
});

els.saveUserBtn.addEventListener("click", async () => {
  if (!can("canManageUsers")) {
    els.userManageStatus.textContent = "Bạn không có quyền quản lý tài khoản.";
    els.userModalStatus.textContent = "Bạn không có quyền quản lý tài khoản.";
    return;
  }

  const userCode = els.userCode.value.trim().toUpperCase();
  const fullName = els.userFullName.value.trim();
  const username = els.userUsername.value.trim().toLowerCase();
  const password = els.userPassword.value;
  const roleKey = els.userRoleKey.value;
  const department = els.userDepartment.value;
  const phone = els.userPhone.value.trim();
  const email = els.userEmail.value.trim();
  const address = els.userAddress.value.trim();
  const bankAccount = els.userBankAccount.value.trim();

  const showFormError = (msg) => {
    els.userManageStatus.textContent = msg;
    els.userModalStatus.textContent = msg;
  };

  try {
    await syncUsersFromRemote(false);
  } catch (err) {
    showToast(`Không thể tải users cloud trước khi lưu: ${err.message}`, "warning");
  }

  if (!userCode) {
    showFormError("Vui lòng nhập mã nhân viên.");
    return;
  }
  if (!fullName || !username || !password || !roleKey) {
    showFormError("Vui lòng nhập đủ họ tên, username, mật khẩu và vai trò.");
    return;
  }

  if (password.length < 6) {
    showFormError("Mật khẩu phải có ít nhất 6 ký tự.");
    return;
  }

  const duplicated = users.find((u) => u.username.toLowerCase() === username && u.id !== editingUserId);
  if (duplicated) {
    showFormError("Username đã tồn tại.");
    return;
  }
  const dupCode = users.find((u) => (u.userCode || "").toUpperCase() === userCode && u.id !== editingUserId);
  if (dupCode) {
    showFormError("Mã nhân viên đã tồn tại.");
    return;
  }

  if (editingUserId) {
    users = users.map((u) => {
      if (u.id !== editingUserId) return u;
      return {
        ...u,
        userCode,
        fullName,
        username,
        password,
        roleKey,
        department,
        phone,
        email,
        address,
        bankAccount
      };
    });
    logActivity("Nhân sự", "Cập nhật tài khoản", `${userCode} | ${username} | Vai trò: ${roleKey}`);
  } else {
    users.push({
      id: `u-${Date.now()}`,
      userCode,
      fullName,
      username,
      password,
      roleKey,
      department,
      phone,
      email,
      address,
      bankAccount,
      status: "active",
      createdAt: Date.now()
    });
    logActivity("Nhân sự", "Tạo tài khoản", `${userCode} | ${username} | Vai trò: ${roleKey}`);
  }

  await persistUsersToRemote(`Lưu tài khoản ${username}`);

  const currentUser = getCurrentUser();
  if (currentUser && currentUser.id === authState.userId) {
    authState.role = getRolePermissions(currentUser.roleKey).label;
    authState.username = currentUser.username;
    saveJSON(STORAGE.auth, authState);
  }

  closeUserModal();
  showToast("Đã lưu thông tin nhân viên.");
  renderAll();
});

els.userBody.addEventListener("click", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  if (!can("canManageUsers")) return;

  const userId = target.dataset.userId;
  if (!userId) return;

  if (target.classList.contains("user-edit-btn")) {
    // close dropdown
    document.querySelectorAll(".user-action-menu").forEach((m) => m.classList.add("hidden"));
    const user = users.find((u) => u.id === userId);
    if (!user) return;
    editingUserId = user.id;
    els.userModalTitle.textContent = `Chỉnh sửa: ${user.fullName}`;
    els.saveUserBtn.textContent = "Cập nhật tài khoản";
    els.userCode.value = user.userCode || "";
    els.userFullName.value = user.fullName;
    els.userUsername.value = user.username;
    els.userPassword.value = user.password;
    els.userRoleKey.value = user.roleKey;
    els.userDepartment.value = user.department || "Ban điều hành";
    els.userPhone.value = user.phone || "";
    els.userEmail.value = user.email || "";
    els.userAddress.value = user.address || "";
    els.userBankAccount.value = user.bankAccount || "";
    els.userManageStatus.textContent = "";
    els.userModalStatus.textContent = "";
    // auto-assign code if the user was created before userCode feature
    if (!els.userCode.value) els.userCode.value = generateUserCode();
    els.userModal.classList.remove("hidden");
    return;
  }

  if (target.classList.contains("user-action-toggle")) {
    // close all other open menus first
    const menu = document.querySelector(`.user-action-menu[data-user-id="${userId}"]`);
    if (menu instanceof HTMLElement) openActionMenuAtToggle(target, menu);
    return;
  }

  if (target.classList.contains("user-suspend-btn")) {
    // close dropdown
    hideAllActionMenus();
    try {
      await syncUsersFromRemote(false);
    } catch (err) {
      showToast(`Không thể tải users cloud trước khi cập nhật: ${err.message}`, "warning");
    }
    const isSelf = userId === authState.userId;
    users = users.map((u) => {
      if (u.id !== userId) return u;
      const newStatus = (u.status || "active") === "suspended" ? "active" : "suspended";
      logActivity("Nhân sự", newStatus === "suspended" ? "Tạm dừng tài khoản" : "Kích hoạt tài khoản", `${u.userCode || u.username}`);
      return { ...u, status: newStatus };
    });
    await persistUsersToRemote("Cập nhật trạng thái tài khoản");
    if (isSelf) {
      performLogout();
      return;
    }
    renderUserTable();
    return;
  }

  if (target.classList.contains("user-delete-btn")) {
    // close dropdown
    hideAllActionMenus();
    if (userId === authState.userId) {
      els.userManageStatus.textContent = "Không thể xóa tài khoản đang đăng nhập.";
      return;
    }

    try {
      await syncUsersFromRemote(false);
    } catch (err) {
      showToast(`Không thể tải users cloud trước khi xóa: ${err.message}`, "warning");
    }

    const deletedUser = users.find((u) => u.id === userId);
    users = users.filter((u) => u.id !== userId);
    await persistUsersToRemote(`Xóa tài khoản ${deletedUser ? deletedUser.username : userId}`);
    els.userManageStatus.textContent = "Đã xóa tài khoản.";
    if (deletedUser) {
      const restoreRef = addToRecycleBin("user", deletedUser, `Tài khoản: ${deletedUser.username}`);
      logActivity("Nhân sự", "Xóa tài khoản", `Username: ${deletedUser.username}`, { restoreRef });
    }
    renderAll();
  }
});

applyDataSourceConfigToInputs();
rememberUsersSyncEndpointFromSource();
syncUsersFromRemote(false).then((updated) => {
  if (updated && authState.loggedIn) renderUserTable();
}).catch(() => {
  // Keep local users as fallback when remote is unavailable.
});
syncCriticalStateFromRemote(false).then((updated) => {
  if (updated) renderAll();
}).catch(() => {
  // Keep local critical state as fallback when remote is unavailable.
});
startUsersAutoSync();
startCriticalStateAutoSync();

if (els.sourceType) {
  els.sourceType.addEventListener("change", () => {
    saveDataSourceConfigFromInputs();
    rememberUsersSyncEndpointFromSource();
  });
}

if (els.sourceUrl) {
  els.sourceUrl.addEventListener("change", () => {
    saveDataSourceConfigFromInputs();
    rememberUsersSyncEndpointFromSource();
  });
}

els.syncBtn.addEventListener("click", async () => {
  if (!can("canSyncData")) {
    els.submitStatus.textContent = "Bạn không có quyền đồng bộ dữ liệu.";
    return;
  }

  const type = els.sourceType.value;
  const url = els.sourceUrl.value.trim();
  saveDataSourceConfigFromInputs();
  rememberUsersSyncEndpointFromSource();
  showToast("Đang đồng bộ dữ liệu...", "info");

  try {
    if (type === "local") {
      els.submitStatus.textContent = "Đang dùng dữ liệu local.";
      return;
    }

    const remoteReports = await fetchRemoteReports(type, url);
    reports = remoteReports;
    saveJSON(STORAGE.reports, reports);
    if (type === "api") {
      try {
        await syncUsersFromRemote(false);
      } catch {
        // reports sync can still succeed even if users endpoint is unavailable.
      }
    }
    renderAll();
    showToast("Đồng bộ thành công.");
    logActivity("Dữ liệu", "Đồng bộ dữ liệu", `Nguồn: ${type} | Endpoint: ${url}`);
  } catch (err) {
    showToast(`Lỗi đồng bộ: ${err.message}`, "error");
  }
});

els.testConnectionBtn.addEventListener("click", async () => {
  const type = els.sourceType.value;
  const url = els.sourceUrl.value.trim();
  saveDataSourceConfigFromInputs();
  rememberUsersSyncEndpointFromSource();
  if (!url) {
    showToast("Vui lòng nhập endpoint URL.", "warning");
    return;
  }

  showToast("Đang kiểm tra kết nối...", "info");

  try {
    if (type === "local") {
      showToast("Dữ liệu local luôn khả dụng.", "success");
      return;
    }

    const response = await fetch(url, { method: "GET" });
    if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);

    if (type === "sheet") {
      const text = await response.text();
      const parsed = parseCsvText(text);
      showToast(`Kết nối thành công. Tìm thấy ${parsed.length} bản ghi CSV.`, "success");
    } else {
      const data = await response.json();
      const count = Array.isArray(data) ? data.length : "N/A";
      showToast(`Kết nối thành công. Dữ liệu: ${count} bản ghi.`, "success");
    }
  } catch (err) {
    showToast(`Lỗi kết nối: ${err.message}`, "error");
  }
});

// ─── Populate Telegram inputs from saved config ─────────────────────────────
function populateTelegramInputs() {
  if (els.telegramBotToken) els.telegramBotToken.value = telegramSourceConfig.token || "";
  if (els.telegramChatId) els.telegramChatId.value = telegramSourceConfig.chatId || "";
  if (els.telegramWebhookBaseUrl) els.telegramWebhookBaseUrl.value = telegramSourceConfig.webhookBaseUrl || "";
  if (els.telegramSyncStatus) {
    if (telegramSourceConfig.lastSyncedAt) {
      const d = new Date(telegramSourceConfig.lastSyncedAt);
      els.telegramSyncStatus.textContent = `Đồng bộ lần cuối: ${d.toLocaleString("vi-VN")}`;
    } else {
      els.telegramSyncStatus.textContent = "Chưa đồng bộ lần nào.";
    }
  }
}
populateTelegramInputs();

function saveTelegramInputs() {
  if (els.telegramBotToken) telegramSourceConfig.token = els.telegramBotToken.value.trim();
  if (els.telegramChatId) telegramSourceConfig.chatId = els.telegramChatId.value.trim();
  if (els.telegramWebhookBaseUrl) telegramSourceConfig.webhookBaseUrl = els.telegramWebhookBaseUrl.value.trim();
  saveTelegramSourceConfig();
}

function stopTelegramRealtimeSync() {
  if (telegramRealtimeSyncTimer) {
    clearInterval(telegramRealtimeSyncTimer);
    telegramRealtimeSyncTimer = null;
  }
}

function startTelegramRealtimeSync() {
  stopTelegramRealtimeSync();
  (async () => {
    if (!authState.loggedIn) return;
    const result = await runTelegramRealtimeSync(true, { fullSync: true });
    if (result && result.importedRows > 0 && els.telegramSyncStatus) {
      els.telegramSyncStatus.textContent = `✓ Đã nhập ${result.importedRows} ca từ queue Telegram (${new Date().toLocaleTimeString("vi-VN")})`;
    }
  })();
  telegramRealtimeSyncTimer = setInterval(async () => {
    if (!authState.loggedIn) return;
    const result = await runTelegramRealtimeSync(true);
    if (result && result.importedRows > 0 && els.telegramSyncStatus) {
      els.telegramSyncStatus.textContent = `✓ Realtime: đã nhập ${result.importedRows} ca mới (${new Date().toLocaleTimeString("vi-VN")})`;
    }
  }, 15000);
}

els.syncTelegramBtn.addEventListener("click", async () => {
  if (!can("canSyncData")) { showToast("Bạn không có quyền đồng bộ dữ liệu.", "warning"); return; }
  saveTelegramInputs();
  if (!telegramSourceConfig.token || !telegramSourceConfig.chatId) {
    showToast("Vui lòng nhập Bot Token và Chat ID.", "warning");
    return;
  }
  els.syncTelegramBtn.disabled = true;
  els.telegramSyncStatus.textContent = "Đang đồng bộ dữ liệu Telegram...";
  try {
    let result;
    try {
      result = await runTelegramRealtimeSync(false, { fullSync: true });
    } catch (bridgeErr) {
      // Fallback: direct Telegram polling if bridge server is not available.
      const direct = await fetchAndParseTelegramMessagesDirect();
      result = { importedRows: direct.importedRows, fetchedRows: direct.parsedMessages, pendingCount: 0, configured: false };
    }
    const msg = `Đã nhập ${result.importedRows} ca mới (${result.fetchedRows} bản ghi nhận về).`;
    els.telegramSyncStatus.textContent = `✓ ${msg} (${new Date().toLocaleTimeString("vi-VN")})`;
    showToast(msg, result.importedRows > 0 ? "success" : "info");
    if (result.importedRows > 0) renderAll();
  } catch (err) {
    els.telegramSyncStatus.textContent = `Lỗi: ${err.message}`;
    showToast(`Lỗi Telegram: ${err.message}`, "error");
  } finally {
    els.syncTelegramBtn.disabled = false;
  }
});

els.testTelegramBtn.addEventListener("click", async () => {
  saveTelegramInputs();
  if (!telegramSourceConfig.token || !telegramSourceConfig.chatId) {
    showToast("Vui lòng nhập Bot Token và Chat ID.", "warning");
    return;
  }
  showToast("Đang cấu hình webhook realtime...", "info");
  try {
    const data = await configureTelegramRealtimeWebhook();
    const webhookUrl = data.webhookUrl || "(chưa đăng ký webhook, thiếu Public URL)";
    const pending = Number(data.pendingCount || 0);
    els.telegramSyncStatus.textContent = `Webhook sẵn sàng. Queue hiện tại: ${pending}.`;
    showToast(`Realtime đã bật. Webhook: ${webhookUrl}`, "success");
    startTelegramRealtimeSync();
    await runTelegramRealtimeSync(false, { fullSync: true });
  } catch (err) {
    showToast(`Lỗi Telegram: ${err.message}`, "error");
  }
});

startTelegramRealtimeSync();

loadRuntimeUsersSyncConfig().then(() => {
  ensureDurableCloudStorage(true).catch(() => {
    // Keep local fallback when cloud storage status cannot be verified.
  });
  syncUsersFromRemote(false).then((updated) => {
    if (updated && authState.loggedIn) renderUserTable();
  }).catch(() => {
    // Keep local users as fallback when runtime endpoint is unavailable.
  });
});

els.reportDate.value = today;
els.reportForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (!can("canSubmitReport")) {
    els.submitStatus.textContent = "Bạn không có quyền nộp báo cáo.";
    return;
  }

  const formData = new FormData(els.reportForm);
  const newRow = {
    date: formData.get("reportDate") || els.reportDate.value,
    department: formData.get("department") || document.querySelector("#department").value,
    submitter: formData.get("submitter") || document.querySelector("#submitter").value,
    completion: Number(formData.get("completion") || document.querySelector("#completion").value),
    quality: Number(formData.get("quality") || document.querySelector("#quality").value),
    issues: Number(formData.get("issues") || document.querySelector("#issues").value),
    updatedAt: Date.now()
  };

  reports.push(newRow);
  saveJSON(STORAGE.reports, reports);
  logActivity("Quy trình", "Nộp báo cáo", `${newRow.department} | ${newRow.submitter} | ${newRow.date}`);

  try {
    const type = els.sourceType.value;
    const url = els.sourceUrl.value.trim();
    saveDataSourceConfigFromInputs();
    if (type !== "local") await postRemoteReport(type, url, newRow);
    showToast("Nộp báo cáo thành công.");
  } catch (err) {
    showToast(`Lưu local thành công, nhưng gửi remote thất bại: ${err.message}`, "warning");
  }

  els.reportForm.reset();
  els.reportDate.value = today;
  renderAll();
});

els.pdfBtn.addEventListener("click", async () => {
  if (!can("canExportPdf")) {
    els.submitStatus.textContent = "Bạn không có quyền xuất PDF.";
    return;
  }

  showToast("Đang tạo PDF...", "info");
  const target = document.querySelector("#dashboardRoot");

  const canvas = await html2canvas(target, { scale: 2, useCORS: true });
  const image = canvas.toDataURL("image/png");
  const pdf = new jsPDF("p", "mm", "a4");
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = (canvas.height * pageWidth) / canvas.width;

  pdf.addImage(image, "PNG", 0, 0, pageWidth, pageHeight);
  pdf.save(`bao-cao-kpi-${today}.pdf`);
  showToast("Đã xuất PDF thành công.");
});

initBrandLogo();
applyPresetFilter("today");
const pageFromLocation = getPageKeyFromLocation();
if (pageFromLocation) activePage = pageFromLocation;
renderAll();
if (authState.loggedIn) {
  setActivePage(activePage, { fromHistory: true, syncUrl: true, replaceUrl: true });
}
applyLoginPrefsToForm();
updateBackButtonState();
