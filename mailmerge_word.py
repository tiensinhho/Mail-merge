from mailmerge import MailMerge
from datetime import date
import os

# 1. Định nghĩa file mẫu (template) và dữ liệu
template_file = 'template.docx'
output_dir = 'Output'
data = [
    {'first_name': 'Nguyễn Văn A', 'last_name': '1.000.000'},
    {'first_name': 'Trần Thị B', 'last_name': '2.500.000'}
]

# Tạo thư mục đầu ra nếu chưa có
os.makedirs(output_dir, exist_ok=True)

for index, record in enumerate(data):
    # Mở template
    with MailMerge(template_file) as document:
        # Trộn dữ liệu
        document.merge(
            first_name=record['first_name'],
            last_name=record['last_name']
        )

        # Lưu file đầu ra
        output_filename = os.path.join(
            output_dir, f"Thu_Moi_{index + 1}_{record['first_name']}.docx")
        document.write(output_filename)

    print(f"Đã tạo: {output_filename}")

print("---")
print("Hoàn tất Mail Merge.")
