# ASUS Credit Note PDF Extractor

📄 Công cụ trích xuất dữ liệu từ các file PDF Credit Note của ASUS sang Excel.

## 🚀 Tính năng

- Upload nhiều file PDF cùng lúc
- Tự động trích xuất thông tin sản phẩm
- Liên kết dữ liệu REBATE
- Xuất Excel với format chuyên nghiệp
- Merge cells tự động

## 📋 Các cột dữ liệu

| Cột | Mô tả |
|-----|-------|
| Tên file PDF | Tên file gốc |
| Product | Tên sản phẩm |
| Product line | Credit Note Remark |
| Serial | Serial Number |
| Part No | Mã Part |
| FOB | Giá USD |
| CN FOB | Credit Note Number |
| CN Landing | CN từ REBATE file |
| Landing cost | Chi phí từ REBATE |

## 🛠️ Cài đặt local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 📦 Deploy lên Streamlit Cloud

1. Push code lên GitHub
2. Truy cập [share.streamlit.io](https://share.streamlit.io)
3. Connect với GitHub repo
4. Chọn `app.py` làm main file
5. Deploy!

## 📝 License

MIT License
