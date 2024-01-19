
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 05/28/2023 01:40:15
-- Generated from EDMX file: C:\Users\USER\Desktop\DA\SOURCE\QLKS\QLKS\Models\dataModel.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [dataQLKS];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

-- --------------------------------------------------
-- Dropping existing FOREIGN KEY constraints
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[fk_ma_cv]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[tblNhanVien] DROP CONSTRAINT [fk_ma_cv];
GO
IF OBJECT_ID(N'[dbo].[fk_ma_dv]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[tblDichVuDaDat] DROP CONSTRAINT [fk_ma_dv];
GO
IF OBJECT_ID(N'[dbo].[fk_ma_hd]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[tblDichVuDaDat] DROP CONSTRAINT [fk_ma_hd];
GO
IF OBJECT_ID(N'[dbo].[fk_ma_lp]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[tblPhong] DROP CONSTRAINT [fk_ma_lp];
GO
IF OBJECT_ID(N'[dbo].[fk_ma_nv]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[tblHoaDon] DROP CONSTRAINT [fk_ma_nv];
GO
IF OBJECT_ID(N'[dbo].[fk_ma_pdp]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[tblHoaDon] DROP CONSTRAINT [fk_ma_pdp];
GO
IF OBJECT_ID(N'[dbo].[fk_ma_tang]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[tblPhong] DROP CONSTRAINT [fk_ma_tang];
GO
IF OBJECT_ID(N'[dbo].[fk_ma_tt_2]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[tblPhong] DROP CONSTRAINT [fk_ma_tt_2];
GO
IF OBJECT_ID(N'[dbo].[fk_ma_tthd]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[tblHoaDon] DROP CONSTRAINT [fk_ma_tthd];
GO
IF OBJECT_ID(N'[dbo].[fk_tgd_ma_kh2]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[tblPhieuDatPhong] DROP CONSTRAINT [fk_tgd_ma_kh2];
GO
IF OBJECT_ID(N'[dbo].[fk_tgd_ma_phong_2]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[tblPhieuDatPhong] DROP CONSTRAINT [fk_tgd_ma_phong_2];
GO
IF OBJECT_ID(N'[dbo].[fk_tgd_tt_2]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[tblPhieuDatPhong] DROP CONSTRAINT [fk_tgd_tt_2];
GO
IF OBJECT_ID(N'[dbo].[fk_tin_nhan]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[tblTinNhan] DROP CONSTRAINT [fk_tin_nhan];
GO

-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[tblChucVu]', 'U') IS NOT NULL
    DROP TABLE [dbo].[tblChucVu];
GO
IF OBJECT_ID(N'[dbo].[tblDichVu]', 'U') IS NOT NULL
    DROP TABLE [dbo].[tblDichVu];
GO
IF OBJECT_ID(N'[dbo].[tblDichVuDaDat]', 'U') IS NOT NULL
    DROP TABLE [dbo].[tblDichVuDaDat];
GO
IF OBJECT_ID(N'[dbo].[tblHoaDon]', 'U') IS NOT NULL
    DROP TABLE [dbo].[tblHoaDon];
GO
IF OBJECT_ID(N'[dbo].[tblKhachHang]', 'U') IS NOT NULL
    DROP TABLE [dbo].[tblKhachHang];
GO
IF OBJECT_ID(N'[dbo].[tblLoaiPhong]', 'U') IS NOT NULL
    DROP TABLE [dbo].[tblLoaiPhong];
GO
IF OBJECT_ID(N'[dbo].[tblNhanVien]', 'U') IS NOT NULL
    DROP TABLE [dbo].[tblNhanVien];
GO
IF OBJECT_ID(N'[dbo].[tblPhieuDatPhong]', 'U') IS NOT NULL
    DROP TABLE [dbo].[tblPhieuDatPhong];
GO
IF OBJECT_ID(N'[dbo].[tblPhong]', 'U') IS NOT NULL
    DROP TABLE [dbo].[tblPhong];
GO
IF OBJECT_ID(N'[dbo].[tblTang]', 'U') IS NOT NULL
    DROP TABLE [dbo].[tblTang];
GO
IF OBJECT_ID(N'[dbo].[tblTinhTrangHoaDon]', 'U') IS NOT NULL
    DROP TABLE [dbo].[tblTinhTrangHoaDon];
GO
IF OBJECT_ID(N'[dbo].[tblTinhTrangPhieuDatPhong]', 'U') IS NOT NULL
    DROP TABLE [dbo].[tblTinhTrangPhieuDatPhong];
GO
IF OBJECT_ID(N'[dbo].[tblTinhTrangPhong]', 'U') IS NOT NULL
    DROP TABLE [dbo].[tblTinhTrangPhong];
GO
IF OBJECT_ID(N'[dbo].[tblTinNhan]', 'U') IS NOT NULL
    DROP TABLE [dbo].[tblTinNhan];
GO

-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'tblChucVus'
CREATE TABLE [dbo].[tblChucVus] (
    [ma_chuc_vu] int IDENTITY(1,1) NOT NULL,
    [chuc_vu] nvarchar(50)  NULL
);
GO

-- Creating table 'tblDichVus'
CREATE TABLE [dbo].[tblDichVus] (
    [ma_dv] int IDENTITY(1,1) NOT NULL,
    [ten_dv] nvarchar(100)  NULL,
    [gia] float  NULL,
    [don_vi] nvarchar(30)  NULL,
    [anh] nvarchar(200)  NULL,
    [ton_kho] int  NULL
);
GO

-- Creating table 'tblDichVuDaDats'
CREATE TABLE [dbo].[tblDichVuDaDats] (
    [id] int IDENTITY(1,1) NOT NULL,
    [ma_hd] int  NULL,
    [ma_dv] int  NULL,
    [so_luong] int  NULL
);
GO

-- Creating table 'tblHoaDons'
CREATE TABLE [dbo].[tblHoaDons] (
    [ma_hd] int IDENTITY(1,1) NOT NULL,
    [ma_nv] int  NULL,
    [ma_pdp] int  NULL,
    [ngay_tra_phong] datetime  NULL,
    [ma_tinh_trang] int  NULL,
    [tien_phong] float  NULL,
    [tien_dich_vu] float  NULL,
    [phu_thu] float  NULL,
    [tong_tien] float  NULL
);
GO

-- Creating table 'tblKhachHangs'
CREATE TABLE [dbo].[tblKhachHangs] (
    [ma_kh] nvarchar(50)  NOT NULL,
    [mat_khau] nvarchar(50)  NULL,
    [ho_ten] nvarchar(50)  NULL,
    [cmt] nvarchar(30)  NULL,
    [sdt] nvarchar(15)  NULL,
    [mail] nvarchar(100)  NULL,
    [diem] int  NULL
);
GO

-- Creating table 'tblLoaiPhongs'
CREATE TABLE [dbo].[tblLoaiPhongs] (
    [loai_phong] int IDENTITY(1,1) NOT NULL,
    [mo_ta] nvarchar(50)  NULL,
    [gia] float  NULL,
    [ti_le_phu_thu] int  NULL,
    [anh] nvarchar(300)  NULL
);
GO

-- Creating table 'tblNhanViens'
CREATE TABLE [dbo].[tblNhanViens] (
    [ma_nv] int IDENTITY(1,1) NOT NULL,
    [ho_ten] nvarchar(50)  NULL,
    [ngay_sinh] datetime  NULL,
    [dia_chi] nvarchar(100)  NULL,
    [sdt] nvarchar(15)  NULL,
    [tai_khoan] nvarchar(50)  NULL,
    [mat_khau] nvarchar(50)  NULL,
    [ma_chuc_vu] int  NULL
);
GO

-- Creating table 'tblPhieuDatPhongs'
CREATE TABLE [dbo].[tblPhieuDatPhongs] (
    [ma_pdp] int IDENTITY(1,1) NOT NULL,
    [ma_kh] nvarchar(50)  NULL,
    [ngay_dat] datetime  NULL,
    [ngay_vao] datetime  NULL,
    [ngay_ra] datetime  NULL,
    [ma_phong] int  NULL,
    [thong_tin_khach_thue] nvarchar(400)  NULL,
    [ma_tinh_trang] int  NULL
);
GO

-- Creating table 'tblPhongs'
CREATE TABLE [dbo].[tblPhongs] (
    [ma_phong] int IDENTITY(1,1) NOT NULL,
    [so_phong] nvarchar(10)  NULL,
    [loai_phong] int  NULL,
    [ma_tang] int  NULL,
    [ma_tinh_trang] int  NULL
);
GO

-- Creating table 'tblTangs'
CREATE TABLE [dbo].[tblTangs] (
    [ma_tang] int IDENTITY(1,1) NOT NULL,
    [ten_tang] nvarchar(20)  NULL
);
GO

-- Creating table 'tblTinhTrangHoaDons'
CREATE TABLE [dbo].[tblTinhTrangHoaDons] (
    [ma_tinh_trang] int IDENTITY(1,1) NOT NULL,
    [mo_ta] nvarchar(50)  NULL
);
GO

-- Creating table 'tblTinhTrangPhieuDatPhongs'
CREATE TABLE [dbo].[tblTinhTrangPhieuDatPhongs] (
    [ma_tinh_trang] int IDENTITY(1,1) NOT NULL,
    [tinh_trang] nvarchar(50)  NULL
);
GO

-- Creating table 'tblTinhTrangPhongs'
CREATE TABLE [dbo].[tblTinhTrangPhongs] (
    [ma_tinh_trang] int IDENTITY(1,1) NOT NULL,
    [mo_ta] nvarchar(50)  NULL
);
GO

-- Creating table 'tblTinNhans'
CREATE TABLE [dbo].[tblTinNhans] (
    [id] int IDENTITY(1,1) NOT NULL,
    [ma_kh] nvarchar(50)  NULL,
    [ho_ten] nvarchar(100)  NULL,
    [mail] nvarchar(100)  NULL,
    [noi_dung] nvarchar(500)  NULL,
    [ngay_gui] datetime  NULL,
    [danh_gia] int  NOT NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [ma_chuc_vu] in table 'tblChucVus'
ALTER TABLE [dbo].[tblChucVus]
ADD CONSTRAINT [PK_tblChucVus]
    PRIMARY KEY CLUSTERED ([ma_chuc_vu] ASC);
GO

-- Creating primary key on [ma_dv] in table 'tblDichVus'
ALTER TABLE [dbo].[tblDichVus]
ADD CONSTRAINT [PK_tblDichVus]
    PRIMARY KEY CLUSTERED ([ma_dv] ASC);
GO

-- Creating primary key on [id] in table 'tblDichVuDaDats'
ALTER TABLE [dbo].[tblDichVuDaDats]
ADD CONSTRAINT [PK_tblDichVuDaDats]
    PRIMARY KEY CLUSTERED ([id] ASC);
GO

-- Creating primary key on [ma_hd] in table 'tblHoaDons'
ALTER TABLE [dbo].[tblHoaDons]
ADD CONSTRAINT [PK_tblHoaDons]
    PRIMARY KEY CLUSTERED ([ma_hd] ASC);
GO

-- Creating primary key on [ma_kh] in table 'tblKhachHangs'
ALTER TABLE [dbo].[tblKhachHangs]
ADD CONSTRAINT [PK_tblKhachHangs]
    PRIMARY KEY CLUSTERED ([ma_kh] ASC);
GO

-- Creating primary key on [loai_phong] in table 'tblLoaiPhongs'
ALTER TABLE [dbo].[tblLoaiPhongs]
ADD CONSTRAINT [PK_tblLoaiPhongs]
    PRIMARY KEY CLUSTERED ([loai_phong] ASC);
GO

-- Creating primary key on [ma_nv] in table 'tblNhanViens'
ALTER TABLE [dbo].[tblNhanViens]
ADD CONSTRAINT [PK_tblNhanViens]
    PRIMARY KEY CLUSTERED ([ma_nv] ASC);
GO

-- Creating primary key on [ma_pdp] in table 'tblPhieuDatPhongs'
ALTER TABLE [dbo].[tblPhieuDatPhongs]
ADD CONSTRAINT [PK_tblPhieuDatPhongs]
    PRIMARY KEY CLUSTERED ([ma_pdp] ASC);
GO

-- Creating primary key on [ma_phong] in table 'tblPhongs'
ALTER TABLE [dbo].[tblPhongs]
ADD CONSTRAINT [PK_tblPhongs]
    PRIMARY KEY CLUSTERED ([ma_phong] ASC);
GO

-- Creating primary key on [ma_tang] in table 'tblTangs'
ALTER TABLE [dbo].[tblTangs]
ADD CONSTRAINT [PK_tblTangs]
    PRIMARY KEY CLUSTERED ([ma_tang] ASC);
GO

-- Creating primary key on [ma_tinh_trang] in table 'tblTinhTrangHoaDons'
ALTER TABLE [dbo].[tblTinhTrangHoaDons]
ADD CONSTRAINT [PK_tblTinhTrangHoaDons]
    PRIMARY KEY CLUSTERED ([ma_tinh_trang] ASC);
GO

-- Creating primary key on [ma_tinh_trang] in table 'tblTinhTrangPhieuDatPhongs'
ALTER TABLE [dbo].[tblTinhTrangPhieuDatPhongs]
ADD CONSTRAINT [PK_tblTinhTrangPhieuDatPhongs]
    PRIMARY KEY CLUSTERED ([ma_tinh_trang] ASC);
GO

-- Creating primary key on [ma_tinh_trang] in table 'tblTinhTrangPhongs'
ALTER TABLE [dbo].[tblTinhTrangPhongs]
ADD CONSTRAINT [PK_tblTinhTrangPhongs]
    PRIMARY KEY CLUSTERED ([ma_tinh_trang] ASC);
GO

-- Creating primary key on [id] in table 'tblTinNhans'
ALTER TABLE [dbo].[tblTinNhans]
ADD CONSTRAINT [PK_tblTinNhans]
    PRIMARY KEY CLUSTERED ([id] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- Creating foreign key on [ma_chuc_vu] in table 'tblNhanViens'
ALTER TABLE [dbo].[tblNhanViens]
ADD CONSTRAINT [fk_ma_cv]
    FOREIGN KEY ([ma_chuc_vu])
    REFERENCES [dbo].[tblChucVus]
        ([ma_chuc_vu])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'fk_ma_cv'
CREATE INDEX [IX_fk_ma_cv]
ON [dbo].[tblNhanViens]
    ([ma_chuc_vu]);
GO

-- Creating foreign key on [ma_dv] in table 'tblDichVuDaDats'
ALTER TABLE [dbo].[tblDichVuDaDats]
ADD CONSTRAINT [fk_ma_dv]
    FOREIGN KEY ([ma_dv])
    REFERENCES [dbo].[tblDichVus]
        ([ma_dv])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'fk_ma_dv'
CREATE INDEX [IX_fk_ma_dv]
ON [dbo].[tblDichVuDaDats]
    ([ma_dv]);
GO

-- Creating foreign key on [ma_hd] in table 'tblDichVuDaDats'
ALTER TABLE [dbo].[tblDichVuDaDats]
ADD CONSTRAINT [fk_ma_hd]
    FOREIGN KEY ([ma_hd])
    REFERENCES [dbo].[tblHoaDons]
        ([ma_hd])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'fk_ma_hd'
CREATE INDEX [IX_fk_ma_hd]
ON [dbo].[tblDichVuDaDats]
    ([ma_hd]);
GO

-- Creating foreign key on [ma_nv] in table 'tblHoaDons'
ALTER TABLE [dbo].[tblHoaDons]
ADD CONSTRAINT [fk_ma_nv]
    FOREIGN KEY ([ma_nv])
    REFERENCES [dbo].[tblNhanViens]
        ([ma_nv])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'fk_ma_nv'
CREATE INDEX [IX_fk_ma_nv]
ON [dbo].[tblHoaDons]
    ([ma_nv]);
GO

-- Creating foreign key on [ma_pdp] in table 'tblHoaDons'
ALTER TABLE [dbo].[tblHoaDons]
ADD CONSTRAINT [fk_ma_pdp]
    FOREIGN KEY ([ma_pdp])
    REFERENCES [dbo].[tblPhieuDatPhongs]
        ([ma_pdp])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'fk_ma_pdp'
CREATE INDEX [IX_fk_ma_pdp]
ON [dbo].[tblHoaDons]
    ([ma_pdp]);
GO

-- Creating foreign key on [ma_tinh_trang] in table 'tblHoaDons'
ALTER TABLE [dbo].[tblHoaDons]
ADD CONSTRAINT [fk_ma_tthd]
    FOREIGN KEY ([ma_tinh_trang])
    REFERENCES [dbo].[tblTinhTrangHoaDons]
        ([ma_tinh_trang])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'fk_ma_tthd'
CREATE INDEX [IX_fk_ma_tthd]
ON [dbo].[tblHoaDons]
    ([ma_tinh_trang]);
GO

-- Creating foreign key on [ma_kh] in table 'tblPhieuDatPhongs'
ALTER TABLE [dbo].[tblPhieuDatPhongs]
ADD CONSTRAINT [fk_tgd_ma_kh2]
    FOREIGN KEY ([ma_kh])
    REFERENCES [dbo].[tblKhachHangs]
        ([ma_kh])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'fk_tgd_ma_kh2'
CREATE INDEX [IX_fk_tgd_ma_kh2]
ON [dbo].[tblPhieuDatPhongs]
    ([ma_kh]);
GO

-- Creating foreign key on [ma_kh] in table 'tblTinNhans'
ALTER TABLE [dbo].[tblTinNhans]
ADD CONSTRAINT [fk_tin_nhan]
    FOREIGN KEY ([ma_kh])
    REFERENCES [dbo].[tblKhachHangs]
        ([ma_kh])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'fk_tin_nhan'
CREATE INDEX [IX_fk_tin_nhan]
ON [dbo].[tblTinNhans]
    ([ma_kh]);
GO

-- Creating foreign key on [loai_phong] in table 'tblPhongs'
ALTER TABLE [dbo].[tblPhongs]
ADD CONSTRAINT [fk_ma_lp]
    FOREIGN KEY ([loai_phong])
    REFERENCES [dbo].[tblLoaiPhongs]
        ([loai_phong])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'fk_ma_lp'
CREATE INDEX [IX_fk_ma_lp]
ON [dbo].[tblPhongs]
    ([loai_phong]);
GO

-- Creating foreign key on [ma_phong] in table 'tblPhieuDatPhongs'
ALTER TABLE [dbo].[tblPhieuDatPhongs]
ADD CONSTRAINT [fk_tgd_ma_phong_2]
    FOREIGN KEY ([ma_phong])
    REFERENCES [dbo].[tblPhongs]
        ([ma_phong])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'fk_tgd_ma_phong_2'
CREATE INDEX [IX_fk_tgd_ma_phong_2]
ON [dbo].[tblPhieuDatPhongs]
    ([ma_phong]);
GO

-- Creating foreign key on [ma_tinh_trang] in table 'tblPhieuDatPhongs'
ALTER TABLE [dbo].[tblPhieuDatPhongs]
ADD CONSTRAINT [fk_tgd_tt_2]
    FOREIGN KEY ([ma_tinh_trang])
    REFERENCES [dbo].[tblTinhTrangPhieuDatPhongs]
        ([ma_tinh_trang])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'fk_tgd_tt_2'
CREATE INDEX [IX_fk_tgd_tt_2]
ON [dbo].[tblPhieuDatPhongs]
    ([ma_tinh_trang]);
GO

-- Creating foreign key on [ma_tang] in table 'tblPhongs'
ALTER TABLE [dbo].[tblPhongs]
ADD CONSTRAINT [fk_ma_tang]
    FOREIGN KEY ([ma_tang])
    REFERENCES [dbo].[tblTangs]
        ([ma_tang])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'fk_ma_tang'
CREATE INDEX [IX_fk_ma_tang]
ON [dbo].[tblPhongs]
    ([ma_tang]);
GO

-- Creating foreign key on [ma_tinh_trang] in table 'tblPhongs'
ALTER TABLE [dbo].[tblPhongs]
ADD CONSTRAINT [fk_ma_tt_2]
    FOREIGN KEY ([ma_tinh_trang])
    REFERENCES [dbo].[tblTinhTrangPhongs]
        ([ma_tinh_trang])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'fk_ma_tt_2'
CREATE INDEX [IX_fk_ma_tt_2]
ON [dbo].[tblPhongs]
    ([ma_tinh_trang]);
GO

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------