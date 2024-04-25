import tkinter as tk
from tkinter import filedialog
from PIL import Image
import pytesseract
from docx import Document

# Thiết lập đường dẫn tessdata
tessdata_dir = "D:/Tesseract-OCR/tessdata"

# Thiết lập pytesseract
pytesseract.pytesseract.tesseract_cmd = "D:/Tesseract-OCR/tesseract.exe"
pytesseract.pytesseract.tessdata_dir_config = tessdata_dir

# Hàm xử lý sự kiện khi nhấn nút "Browse"
def browse_image():
    # Hiển thị hộp thoại để chọn file ảnh
    file_path = filedialog.askopenfilename(title="Chọn ảnh", filetypes=(("Image files", "*.png *.jpg *.jpeg"), ("All files", "*.*")))
    
    if file_path:
        # Đọc ảnh bằng thư viện Pillow
        image = Image.open(file_path)
        
        # Chuyển đổi ảnh thành văn bản bằng pytesseract
        text = pytesseract.image_to_string(image, lang="vie", config="--oem 3 --psm 6")
        
        # Hiển thị kết quả
        result_text.delete('1.0', tk.END)
        result_text.insert(tk.END, text)

# Hàm xử lý sự kiện khi nhấn nút "Export"
def export_text():
    # Hiển thị hộp thoại để chọn vị trí và tên file để lưu kết quả
    file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=(("Word files", "*.docx"), ("All files", "*.*")))
    
    if file_path:
        # Lấy nội dung từ ô văn bản
        text = result_text.get("1.0", tk.END)
        
        # Tạo một tài liệu Word mới
        doc = Document()
        
        # Thêm nội dung vào tài liệu
        doc.add_paragraph(text)
        
        # Lưu tài liệu vào file
        doc.save(file_path)
            
        # Hiển thị thông báo lưu thành công
        tk.messagebox.showinfo("Export", "Lưu kết quả thành công.")

# Tạo cửa sổ giao diện
window = tk.Tk()
window.title("Chuyển đổi ảnh thành văn bản")
window.geometry("400x300")

# Tạo nút "Browse" để chọn ảnh
browse_button = tk.Button(window, text="Chọn ảnh", command=browse_image)
browse_button.pack(pady=10)

# Tạo ô văn bản để hiển thị kết quả
result_text = tk.Text(window, height=10, width=40)
result_text.pack()

# Tạo nút "Export" để xuất kết quả
export_button = tk.Button(window, text="Xuất file", command=export_text)
export_button.pack(pady=10)

# Chạy chương trình
window.mainloop()
