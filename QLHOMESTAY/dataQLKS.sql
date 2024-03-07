use master
if exists (select*from sysdatabases where name = 'dataQLKS')
drop database dataQLKS
go 
create database dataQLKS
go
use dataQLKS
go

CREATE TABLE tblChucVu (
    ma_chuc_vu int IDENTITY(1,1) NOT NULL,
    chuc_vu nvarchar(50)  NULL,
	PRIMARY KEY (ma_chuc_vu)
);
GO

CREATE TABLE tblDichVu (
    ma_dv int IDENTITY(1,1) NOT NULL,
    ten_dv nvarchar(100)  NULL,
    gia float  NULL,
    don_vi nvarchar(30)  NULL,
    anh nvarchar(200)  NULL,
    ton_kho int  NULL,
	PRIMARY KEY (ma_dv)

);
GO


CREATE TABLE tblKhachHang (
    ma_kh nvarchar(50)  NOT NULL,
    mat_khau nvarchar(50)  NULL,
    ho_ten nvarchar(50)  NULL,
    cmt nvarchar(30)  NULL,
    sdt nvarchar(15)  NULL,
    mail nvarchar(100)  NULL,
    diem int  NULL,
	PRIMARY KEY (ma_kh)
);
GO

CREATE TABLE tblTinNhan (
    id int IDENTITY(1,1) NOT NULL,
    ma_kh nvarchar(50)  NULL,
    ho_ten nvarchar(100)  NULL,
    mail nvarchar(100)  NULL,
    noi_dung nvarchar(500)  NULL,
    ngay_gui datetime  NULL,
    danh_gia int  NOT NULL,
	PRIMARY KEY (id),
	FOREIGN KEY (ma_kh) REFERENCES tblKhachHang(ma_kh)
);
GO

CREATE TABLE tblLoaiPhong (
    loai_phong int IDENTITY(1,1) NOT NULL,
    mo_ta nvarchar(50)  NULL,
    gia float  NULL,
    ti_le_phu_thu int  NULL,
    anh nvarchar(300)  NULL,
	PRIMARY KEY (loai_phong)
);
GO

CREATE TABLE tblTang (
    ma_tang int IDENTITY(1,1) NOT NULL,
    ten_tang nvarchar(20)  NULL,
	PRIMARY KEY (ma_tang)
);
GO

CREATE TABLE tblTinhTrangHoaDon (
    ma_tinh_trang int IDENTITY(1,1) NOT NULL,
    mo_ta nvarchar(50)  NULL,
	PRIMARY KEY (ma_tinh_trang)
);
GO

CREATE TABLE tblTinhTrangPhieuDatPhong (
    ma_tinh_trang int IDENTITY(1,1) NOT NULL,
    tinh_trang nvarchar(50)  NULL,
	PRIMARY KEY (ma_tinh_trang)
);
GO

CREATE TABLE tblTinhTrangPhong (
    ma_tinh_trang int IDENTITY(1,1) NOT NULL,
    mo_ta nvarchar(50)  NULL,
	PRIMARY KEY (ma_tinh_trang)
);
GO

CREATE TABLE tblPhong (
    ma_phong int IDENTITY(1,1) NOT NULL,
    so_phong nvarchar(10)  NULL,
    loai_phong int  NULL,
    ma_tang int  NULL,
    ma_tinh_trang int  NULL,
	PRIMARY KEY (ma_phong),
	FOREIGN KEY (loai_phong) REFERENCES tblLoaiPhong(loai_phong),
	FOREIGN KEY (ma_tang) REFERENCES tblTang(ma_tang),
	FOREIGN KEY (ma_tinh_trang) REFERENCES tblTinhTrangPhong(ma_tinh_trang)
);
GO

CREATE TABLE tblPhieuDatPhong (
    ma_pdp int IDENTITY(1,1) NOT NULL,
    ma_kh nvarchar(50)  NULL,
    ngay_dat datetime  NULL,
    ngay_vao datetime  NULL,
    ngay_ra datetime  NULL,
    ma_phong int  NULL,
    thong_tin_khach_thue nvarchar(400)  NULL,
    ma_tinh_trang int  NULL,
	PRIMARY KEY (ma_pdp),
	FOREIGN KEY (ma_kh) REFERENCES tblKhachHang(ma_kh),
	FOREIGN KEY (ma_phong) REFERENCES tblPhong(ma_phong),
	FOREIGN KEY (ma_tinh_trang) REFERENCES tblTinhTrangPhieuDatPhong(ma_tinh_trang)
);
GO

CREATE TABLE tblNhanVien (
    ma_nv int IDENTITY(1,1) NOT NULL,
    ho_ten nvarchar(50)  NULL,
    ngay_sinh datetime  NULL,
    dia_chi nvarchar(100)  NULL,
    sdt nvarchar(15)  NULL,
    tai_khoan nvarchar(50)  NULL,
    mat_khau nvarchar(50)  NULL,
    ma_chuc_vu int  NULL,
	PRIMARY KEY (ma_nv),
	FOREIGN KEY (ma_chuc_vu) REFERENCES tblChucVu(ma_chuc_vu)
);
GO

CREATE TABLE tblHoaDon (
    ma_hd int IDENTITY(1,1) NOT NULL,
    ma_nv int  NULL,
    ma_pdp int  NULL,
    ngay_tra_phong datetime  NULL,
    ma_tinh_trang int  NULL,
    tien_phong float  NULL,
    tien_dich_vu float  NULL,
    phu_thu float  NULL,
    tong_tien float  NULL,
	PRIMARY KEY (ma_hd),
	FOREIGN KEY (ma_nv) REFERENCES tblNhanVien(ma_nv),
	FOREIGN KEY (ma_pdp) REFERENCES tblPhieuDatPhong(ma_pdp),
	FOREIGN KEY (ma_tinh_trang) REFERENCES tblTinhTrangHoaDon(ma_tinh_trang)
);

GO
CREATE TABLE tblDichVuDaDat (
    id int IDENTITY(1,1) NOT NULL,
    ma_hd int  NULL,
    ma_dv int  NULL,
    so_luong int  NULL,
	PRIMARY KEY (id),
	FOREIGN KEY (ma_dv) REFERENCES tblDichVu(ma_dv),
	FOREIGN KEY (ma_hd) REFERENCES tblHoaDon(ma_hd)
);
GO


insert into tblChucVu values(N'Quản trị viên')
insert into tblChucVu values(N'Quản lý')
insert into tblChucVu values(N'Nhân Viên')
select * from tblChucVu

insert into tblDichVu values(N'Thuê Xe Máy (xe số)',150000,N'1xe/ ngày','/Content/Images/DichVu/xe1.jpg',15)
insert into tblDichVu values(N'Thuê Xe Máy (xe ga)',2500000,N'1xe/ ngày','/Content/Images/DichVu/xe2.jpg',19)
insert into tblDichVu values(N'Thuê Xe Máy (xe côn)',250000,N'1xe/ ngày','/Content/Images/DichVu/xe3.jpg',20)
insert into tblDichVu values(N'Thuê Xe ôto tự lái',1100000,N'1xe/ ngày','/Content/Images/DichVu/xe4.jpg',5)
insert into tblDichVu values(N'Dụng cụ cắm trại 1',500000,N'1bộ/ ngày','/Content/Images/DichVu/cp1.jpg',19)
insert into tblDichVu values(N'Dụng cụ cắm trại 2',850000,N'Bộ','/Content/Images/DichVu/cp2.jpg',19)
insert into tblDichVu values(N'Combo Nướng',500000,N'Combo','/Content/Images/DichVu/n1.jpg',997)
insert into tblDichVu values(N'Combo Nướng Lẩu',850000,N'Combo','/Content/Images/DichVu/n2.jpg',9998)

insert into tblLoaiPhong values(N'CamPing',850000,10,'["/Content/Images/Phong/c1.jpg"]')
insert into tblLoaiPhong values(N'Phòng đơn',850000,10,'["/Content/Images/Phong/p2.jpg"]')
insert into tblLoaiPhong values(N'Phòng đôi',1100000,15,'["/Content/Images/Phong/p1.jpg"]')


insert into tblNhanVien values(N'Nam đẹp trai','2024-02-28','123 cây chuối quận 12','01177915896','admin','12345',1)
insert into tblNhanVien values(N'Thọ đẹp trai','2024-02-28','123 cây chuối quận 12','01177915896','quanly','12345',2)
insert into tblNhanVien values(N'Tiến Ngu','2024-02-28','123 cây chuối quận 12','01177915896','nhanvien','12345',3)

insert into tblTang values(N'View Thung Lũng')
insert into tblTang values(N'Dành cho CamPing')
insert into tblTang values(N'View Thành Phố')

insert into tblTinhTrangHoaDon values(N'Chưa thanh toán')
insert into tblTinhTrangHoaDon values(N'Đã thanh toán')

insert into tblTinhTrangPhong values(N'Trống')
insert into tblTinhTrangPhong values(N'Đang sử dụng')
insert into tblTinhTrangPhong values(N'Đang dọn')
insert into tblTinhTrangPhong values(N'Đang bảo trì')
insert into tblTinhTrangPhong values(N'Dừng sử dụng')
select*from tblTinhTrangPhong

insert into tblTinhTrangPhieuDatPhong values(N'Đang đặt')
insert into tblTinhTrangPhieuDatPhong values(N'Đã xong')
insert into tblTinhTrangPhieuDatPhong values(N'Đã hủy')
insert into tblTinhTrangPhieuDatPhong values(N'Đã thanh toán')

insert into tblPhong values('P101',1,1,1)
insert into tblPhong values('P102',1,1,1)
insert into tblPhong values('P103',1,3,1)
insert into tblPhong values('P104',1,2,1)
insert into tblPhong values('P105',1,2,1)
insert into tblPhong values('P106',1,3,1)
insert into tblPhong values('P201',2,1,1)
insert into tblPhong values('P202',2,3,1)
insert into tblPhong values('P203',2,1,1)
insert into tblPhong values('P204',2,1,1)
insert into tblPhong values('P205',2,1,1)
insert into tblPhong values('P206',1,3,1)
insert into tblPhong values('P301',3,1,1)
insert into tblPhong values('P302',3,1,1)
insert into tblPhong values('P303',1,1,1)
insert into tblPhong values('P304',3,3,1)
insert into tblPhong values('P305',3,1,1)
insert into tblPhong values('P306',3,3,1)
select*from  tblPhong