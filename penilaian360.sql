-- Membuat database baru (Opsional, lewati jika sudah ada database)
CREATE DATABASE IF NOT EXISTS db_penilaian_360;
USE db_penilaian_360;

-- Membuat tabel penilaian atasan terhadap bawahan
CREATE TABLE penilaian_bawahan (
    nip_bawahan VARCHAR(18) NOT NULL COMMENT 'NIP tanpa spasi (Primary Key)',
    nama_bawahan VARCHAR(150) NOT NULL,
    jenis_pegawai VARCHAR(50) NOT NULL COMMENT 'Contoh: PNS, PPPK',
    nama_satuan_kerja VARCHAR(200) NOT NULL,
    tahun_penilaian YEAR DEFAULT '2025',
    nip_atasan_penilai VARCHAR(18) NOT NULL COMMENT 'NIP atasan yang menilai',
    waktu_pengisian TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (nip_bawahan)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;