# BÀI LÀM: TRA CỨU DỮ LIỆU EXCEL BẰNG HÀM VLOOKUP

## PHẦN 1. PHÂN TÍCH BÀI TOÁN (I/O)

### 1. Xác định các thành phần của hàm VLOOKUP

| Thành phần | Cách xác định | Ý nghĩa |
|---|---|---|
| `Lookup_value` | `C2` trong sheet `DanhSachBaoCao` | Mã tác giả cần tra cứu, ví dụ `TG001`. |
| `Table_array` | `'DanhMucTacGia'!$A$2:$E$100` | Vùng dữ liệu tra cứu từ cột A đến cột E của sheet `DanhMucTacGia`. Dấu `$` giúp cố định vùng khi kéo công thức. |
| `Col_index_num` cho cột D | `2` | Trả về tên tác giả từ cột B trong vùng tra cứu. |
| `Col_index_num` cho cột E | `4` | Trả về phòng trình bày từ cột D trong vùng tra cứu. |
| `Col_index_num` cho cột F | `5` | Trả về thù lao từ cột E trong vùng tra cứu. |

**Lưu ý:** Ví dụ trên giả sử danh mục tác giả có dữ liệu từ hàng 2 đến hàng 100. Cần điều chỉnh hàng cuối theo dữ liệu thực tế. Vùng tra cứu phải bắt đầu từ cột A vì cột A chứa Mã tác giả dùng để tìm kiếm.

### 2. Giải pháp xử lý 3 bẫy dữ liệu

- **Bẫy 1 – Mã tác giả có khoảng trắng thừa:** Sử dụng hàm `TRIM(C2)` để loại bỏ khoảng trắng thừa ở đầu và cuối mã tác giả trước khi tra cứu.

- **Bẫy 2 – Mã tác giả chưa có trong danh mục:** Sử dụng hàm `IFERROR` bao quanh `VLOOKUP` để hiển thị `"Chưa cập nhật"` thay vì lỗi `#N/A`.

- **Bẫy 3 – Vùng tra cứu bị trôi khi kéo công thức:** Sử dụng tham chiếu tuyệt đối `$A$2:$E$100` để cố định vùng dữ liệu tra cứu.

Cấu trúc công thức tổng quát:

```excel
=IFERROR(VLOOKUP(TRIM(C2),'DanhMucTacGia'!$A$2:$E$100,2,FALSE),"Chưa cập nhật")