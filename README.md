# Quản lý Trung tâm Nhật Như — Cloud

Website React/Vite. Dữ liệu nghiệp vụ được lưu localStorage và tự động sao lưu lên Supabase theo tài khoản đăng nhập.

## 1. Tạo Supabase
1. Tạo project tại https://supabase.com/
2. Authentication > Providers > Email: bật Email/Password.
3. SQL Editor: chạy `supabase-schema.sql`.
4. Authentication > Users: tạo tài khoản quản trị bằng email + password.

## 2. Cấu hình local
Copy `.env.example` thành `.env` rồi điền:
- `VITE_SUPABASE_URL`
- `VITE_SUPABASE_ANON_KEY`

## 3. Chạy
```bash
npm install
npm run dev
```

## 4. Deploy GitHub Pages
Repository mặc định phải tên `quan-ly-trung-tam-nhat-nhu`. Workflow GitHub Actions sẽ build và deploy thư mục `dist`.

Nếu đổi tên repository, sửa `base` trong `vite.config.js` và `path` URL tương ứng.

## 5. Dữ liệu
Mỗi thay đổi được lưu local trước, sau đó debounce khoảng 350ms và upsert vào bảng `app_state`. Dữ liệu được tách theo `user_id`, có RLS nên người dùng chỉ đọc/ghi dữ liệu của chính tài khoản đó.
