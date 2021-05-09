-- phpMyAdmin SQL Dump
-- version 4.7.4
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: 03 Mei 2021 pada 06.30
-- Versi Server: 10.1.29-MariaDB
-- PHP Version: 7.2.0

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET AUTOCOMMIT = 0;
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `pendaftaran`
--

-- --------------------------------------------------------

--
-- Struktur dari tabel `peserta_didik`
--

CREATE TABLE `peserta_didik` (
  `id` int(11) NOT NULL,
  `jenis_pendaftaran` varchar(30) NOT NULL,
  `tgl_msk` date NOT NULL,
  `nis` varchar(100) NOT NULL,
  `no_ujian` varchar(100) NOT NULL,
  `paud` varchar(100) NOT NULL,
  `tk` varchar(100) NOT NULL,
  `skhun` varchar(100) NOT NULL,
  `ijazah` varchar(100) NOT NULL,
  `hobi` varchar(100) NOT NULL,
  `cita2` varchar(100) NOT NULL,
  `nama` varchar(100) NOT NULL,
  `jenis_kelamin` varchar(100) NOT NULL,
  `nisn_sekarang` varchar(100) NOT NULL,
  `nik` varchar(100) NOT NULL,
  `tempat_lahir` varchar(100) NOT NULL,
  `tanggal_lahir` date NOT NULL,
  `agama` varchar(100) NOT NULL,
  `berkebutuhan` varchar(100) NOT NULL,
  `alamat` varchar(100) NOT NULL,
  `rt` varchar(100) NOT NULL,
  `rw` varchar(100) NOT NULL,
  `dusun` varchar(100) NOT NULL,
  `kelurahan` varchar(100) NOT NULL,
  `kecamatan` varchar(100) NOT NULL,
  `kode_pos` varchar(100) NOT NULL,
  `tempat_tinggal` varchar(100) NOT NULL,
  `transportasi` varchar(100) NOT NULL,
  `no_hp` varchar(100) NOT NULL,
  `email` varchar(100) NOT NULL,
  `kps` varchar(100) NOT NULL,
  `no_kps` varchar(100) NOT NULL,
  `kewarganegaraan` varchar(100) NOT NULL,
  `negara` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `peserta_didik`
--

INSERT INTO `peserta_didik` (`id`, `jenis_pendaftaran`, `tgl_msk`, `nis`, `no_ujian`, `paud`, `tk`, `skhun`, `ijazah`, `hobi`, `cita2`, `nama`, `jenis_kelamin`, `nisn_sekarang`, `nik`, `tempat_lahir`, `tanggal_lahir`, `agama`, `berkebutuhan`, `alamat`, `rt`, `rw`, `dusun`, `kelurahan`, `kecamatan`, `kode_pos`, `tempat_tinggal`, `transportasi`, `no_hp`, `email`, `kps`, `no_kps`, `kewarganegaraan`, `negara`) VALUES
(116, '01', '2021-05-04', '123456789', '123456789', 'Ya', 'Ya', '123456789', '123456789', 'a', 'a', 'M. Aldan Adiar Firdaus', 'Laki-laki', '123456789', '123456789', 'surabaya', '2005-06-08', '01', '01', 'Jalan merpati nomor 7', '3', '4', 'Dusun suka maju', 'gubeng', 'gubeng', '60123', '01', '01', '123456789', 'aldanfirdaus30@gmail.com', 'Tidak', 'Tidak', 'Indonesia(WNI)', 'Indonesia');

--
-- Indexes for dumped tables
--

--
-- Indexes for table `peserta_didik`
--
ALTER TABLE `peserta_didik`
  ADD PRIMARY KEY (`id`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `peserta_didik`
--
ALTER TABLE `peserta_didik`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=117;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
