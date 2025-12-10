-- phpMyAdmin SQL Dump
-- version 5.2.0
-- https://www.phpmyadmin.net/
--
-- Host: localhost
-- Waktu pembuatan: 23 Nov 2025 pada 19.00
-- Versi server: 8.0.30
-- Versi PHP: 8.2.21

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `korpraport`
--

-- --------------------------------------------------------

--
-- Struktur dari tabel `gender`
--

CREATE TABLE `gender` (
  `kd_gender` char(1) NOT NULL,
  `gender` varchar(20) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

--
-- Dumping data untuk tabel `gender`
--

INSERT INTO `gender` (`kd_gender`, `gender`) VALUES
('L', 'Laki-laki'),
('P', 'Perempuan');

-- --------------------------------------------------------

--
-- Struktur dari tabel `history_log`
--

CREATE TABLE `history_log` (
  `id` int NOT NULL,
  `user_id` int DEFAULT NULL,
  `aksi` varchar(50) NOT NULL,
  `keterangan` text,
  `waktu` timestamp NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

--
-- Dumping data untuk tabel `history_log`
--

INSERT INTO `history_log` (`id`, `user_id`, `aksi`, `keterangan`, `waktu`) VALUES
(1, 2, 'TAMBAH DATA DIRI', 'User 000001 melengkapi data diri', '2025-10-01 16:09:27'),
(2, 3, 'TAMBAH DATA DIRI', 'User 000002 melengkapi data diri', '2025-10-01 16:16:41'),
(3, 2, 'LOGOUT', 'User 000001 logout', '2025-10-04 13:04:32'),
(4, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-04 13:04:39'),
(5, 2, 'LOGOUT', 'User 000001 logout', '2025-10-04 13:04:42'),
(6, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-04 13:04:49'),
(7, 1, 'LOGOUT', 'User 123456 logout', '2025-10-04 13:10:00'),
(8, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-04 13:10:04'),
(9, 2, 'LOGOUT', 'User 000001 logout', '2025-10-04 13:10:16'),
(10, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-04 13:10:25'),
(11, 1, 'LOGOUT', 'User 123456 logout', '2025-10-04 13:14:19'),
(12, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-04 13:14:23'),
(13, 2, 'UPDATE FOTO', 'User mengupdate foto profil', '2025-10-04 13:15:12'),
(14, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-04 13:16:18'),
(15, 2, 'LOGOUT', 'User 000001 logout', '2025-10-04 13:17:39'),
(16, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-04 13:17:45'),
(17, 1, 'LOGOUT', 'User 123456 logout', '2025-10-04 13:35:38'),
(18, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-04 13:37:07'),
(19, 2, 'LOGOUT', 'User 000001 logout', '2025-10-04 13:37:48'),
(20, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-04 13:39:21'),
(21, 2, 'LOGOUT', 'User 000001 logout', '2025-10-04 13:40:44'),
(22, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-04 13:40:59'),
(23, 2, 'LOGOUT', 'User 000001 logout', '2025-10-04 13:43:18'),
(24, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-04 13:43:31'),
(25, 3, 'LOGIN', 'User 000002 berhasil login', '2025-10-04 18:04:34'),
(26, 3, 'LOGOUT', 'User 000002 logout', '2025-10-04 18:04:57'),
(27, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-04 18:05:21'),
(28, 2, 'LOGOUT', 'User 000001 logout', '2025-10-04 18:06:09'),
(29, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-04 18:06:14'),
(30, 1, 'LOGOUT', 'User 123456 logout', '2025-10-04 18:07:39'),
(31, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-04 18:07:43'),
(32, 2, 'LOGOUT', 'User 000001 logout', '2025-10-04 18:12:33'),
(33, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-04 18:16:18'),
(34, 2, 'LOGOUT', 'User 000001 logout', '2025-10-04 18:21:33'),
(35, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-04 18:21:38'),
(36, 1, 'LOGOUT', 'User 123456 logout', '2025-10-04 18:23:08'),
(37, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-04 18:23:14'),
(38, 2, 'LOGOUT', 'User 000001 logout', '2025-10-04 18:23:41'),
(39, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-04 18:23:45'),
(40, 1, 'LOGOUT', 'User 123456 logout', '2025-10-04 18:28:14'),
(41, 3, 'LOGIN', 'User 000002 berhasil login', '2025-10-04 18:28:18'),
(42, 3, 'TAMBAH DATA DIRI', 'User 000002 melengkapi data diri', '2025-10-04 18:30:37'),
(43, 3, 'LOGOUT', 'User 000002 logout', '2025-10-04 18:33:16'),
(44, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-04 18:33:21'),
(45, 1, 'TAMBAH USER', 'Admin menambahkan user baru dengan NRP: 000003', '2025-10-04 18:33:50'),
(46, 1, 'LOGOUT', 'User 123456 logout', '2025-10-04 18:33:59'),
(47, 4, 'LOGIN', 'User 000003 berhasil login', '2025-10-04 18:34:04'),
(48, 4, 'TAMBAH DATA DIRI', 'User 000003 melengkapi data diri', '2025-10-04 18:36:40'),
(49, 4, 'UPDATE DATA DIRI', 'User 000003 mengupdate data diri', '2025-10-04 18:45:00'),
(50, 4, 'LOGOUT', 'User 000003 logout', '2025-10-04 18:45:15'),
(51, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-04 18:45:21'),
(52, 1, 'LOGOUT', 'User 123456 logout', '2025-10-04 18:46:47'),
(53, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-04 18:46:52'),
(54, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-04 18:48:41'),
(55, 2, 'LOGOUT', 'User 000001 logout', '2025-10-04 18:48:47'),
(56, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-04 18:48:51'),
(57, 3, 'LOGOUT', 'User 000002 logout', '2025-10-06 01:59:08'),
(58, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-06 01:59:15'),
(59, 1, 'UPDATE DATA', 'Admin mengupdate data personel NRP: 000001', '2025-10-06 02:04:39'),
(60, 1, 'LOGOUT', 'User 123456 logout', '2025-10-06 02:06:55'),
(61, 3, 'LOGIN', 'User 000002 berhasil login', '2025-10-06 02:08:20'),
(62, 3, 'LOGOUT', 'User 000002 logout', '2025-10-06 02:09:33'),
(63, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-06 02:09:38'),
(64, 1, 'UPDATE DATA', 'Admin mengupdate data personel NRP: 000003', '2025-10-06 02:10:11'),
(65, 1, 'LOGOUT', 'User 123456 logout', '2025-10-06 02:41:40'),
(66, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-06 02:41:45'),
(67, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-06 02:44:21'),
(68, 2, 'LOGOUT', 'User 000001 logout', '2025-10-06 02:44:36'),
(69, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-06 02:44:42'),
(70, 1, 'LOGOUT', 'User 123456 logout', '2025-10-06 02:56:34'),
(71, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-06 02:56:49'),
(72, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-06 02:57:39'),
(73, 2, 'LOGOUT', 'User 000001 logout', '2025-10-06 03:10:00'),
(74, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-06 03:10:07'),
(75, 1, 'LOGOUT', 'User 123456 logout', '2025-10-06 03:15:27'),
(76, 3, 'LOGIN', 'User 000002 berhasil login', '2025-10-06 03:15:38'),
(77, 3, 'LOGOUT', 'User 000002 logout', '2025-10-06 03:17:16'),
(78, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-06 03:17:23'),
(79, 1, 'LOGOUT', 'User 123456 logout', '2025-10-06 03:25:57'),
(80, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-06 04:02:25'),
(81, 1, 'LOGOUT', 'User 123456 logout', '2025-10-06 04:03:03'),
(82, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-06 04:03:10'),
(83, 2, 'LOGOUT', 'User 000001 logout', '2025-10-06 04:03:23'),
(84, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-06 04:14:59'),
(85, 1, 'UPDATE DATA', 'Admin mengupdate data personel NRP: 000003', '2025-10-06 04:17:24'),
(86, 1, 'LOGOUT', 'User 123456 logout', '2025-10-06 04:17:29'),
(87, 4, 'LOGIN', 'User 000003 berhasil login', '2025-10-06 04:17:37'),
(88, 4, 'UPDATE FOTO', 'User mengupdate foto profil', '2025-10-06 04:18:08'),
(89, 4, 'UPDATE DATA DIRI', 'User 000003 mengupdate data diri', '2025-10-06 04:20:50'),
(90, 4, 'LOGOUT', 'User 000003 logout', '2025-10-06 04:20:52'),
(91, 3, 'LOGIN', 'User 000002 berhasil login', '2025-10-06 04:20:56'),
(92, 3, 'UPDATE DATA DIRI', 'User 000002 mengupdate data diri', '2025-10-06 04:22:32'),
(93, 3, 'UPDATE FOTO', 'User mengupdate foto profil', '2025-10-06 04:24:03'),
(94, 3, 'UPDATE FOTO', 'User mengupdate foto profil', '2025-10-06 04:24:10'),
(95, 3, 'UPDATE FOTO', 'User mengupdate foto profil', '2025-10-06 04:24:50'),
(96, 3, 'UPDATE FOTO', 'User mengupdate foto profil', '2025-10-06 04:24:55'),
(97, 3, 'LOGOUT', 'User 000002 logout', '2025-10-06 04:25:10'),
(98, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-06 04:25:14'),
(99, 1, 'LOGOUT', 'User 123456 logout', '2025-10-06 04:26:30'),
(100, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-06 04:26:37'),
(101, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-09 03:02:29'),
(102, 2, 'LOGOUT', 'User 000001 logout', '2025-10-09 03:03:11'),
(103, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-09 03:03:23'),
(104, 1, 'LOGOUT', 'User 123456 logout', '2025-10-09 03:08:27'),
(105, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-09 03:08:56'),
(106, 1, 'LOGOUT', 'User 123456 logout', '2025-10-09 03:09:01'),
(107, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-09 03:09:06'),
(108, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-09 03:13:27'),
(109, 1, 'LOGOUT', 'User 123456 logout', '2025-10-09 03:18:19'),
(110, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-09 03:22:33'),
(111, 1, 'LOGOUT', 'User 123456 logout', '2025-10-09 03:22:48'),
(112, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-09 03:24:22'),
(113, 2, 'LOGOUT', 'User 000001 logout', '2025-10-09 03:24:40'),
(114, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-20 03:56:06'),
(115, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-20 04:05:42'),
(116, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-20 04:09:22'),
(117, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-20 04:09:58'),
(118, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-20 04:10:03'),
(119, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-20 04:10:30'),
(120, 2, 'LOGOUT', 'User 000001 logout', '2025-10-20 04:10:31'),
(121, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-20 04:10:38'),
(122, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-20 04:19:52'),
(123, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-20 04:20:32'),
(124, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-20 04:20:38'),
(125, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-20 04:24:02'),
(126, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-20 04:24:09'),
(127, 2, 'LOGOUT', 'User 000001 logout', '2025-10-20 04:24:11'),
(128, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-20 04:24:20'),
(129, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-20 04:24:37'),
(130, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-20 04:24:55'),
(131, 2, 'LOGOUT', 'User 000001 logout', '2025-10-20 04:25:42'),
(132, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-23 03:18:24'),
(133, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-23 03:22:43'),
(134, 2, 'LOGOUT', 'User 000001 logout', '2025-10-23 03:22:47'),
(135, 3, 'LOGIN', 'User 000002 berhasil login', '2025-10-23 03:23:15'),
(136, 3, 'UPDATE DATA DIRI', 'User 000002 mengupdate data diri', '2025-10-23 03:23:41'),
(137, 3, 'UPDATE DATA DIRI', 'User 000002 mengupdate data diri', '2025-10-23 03:24:09'),
(138, 3, 'LOGOUT', 'User 000002 logout', '2025-10-23 03:24:46'),
(139, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-23 04:08:56'),
(140, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-23 04:13:40'),
(141, 2, 'UPDATE FOTO', 'User mengupdate foto profil', '2025-10-23 04:14:06'),
(142, 2, 'LOGOUT', 'User 000001 logout', '2025-10-23 04:20:42'),
(143, 3, 'LOGIN', 'User 000002 berhasil login', '2025-10-23 04:21:10'),
(144, 3, 'LOGIN', 'User 000002 berhasil login', '2025-10-27 02:26:26'),
(145, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-27 03:16:14'),
(146, 2, 'LOGOUT', 'User 000001 logout', '2025-10-27 03:16:42'),
(147, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-27 03:16:49'),
(148, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-29 17:16:30'),
(149, 2, 'LOGOUT', 'User 000001 logout', '2025-10-29 17:22:46'),
(150, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-29 17:23:01'),
(151, 1, 'EXPORT DATA', 'Admin export data personel ke Excel', '2025-10-29 17:36:45'),
(152, 1, 'EXPORT DATA', 'Admin export data personel ke Excel', '2025-10-29 18:00:15'),
(153, 1, 'LOGOUT', 'User 123456 logout', '2025-10-29 18:00:50'),
(154, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-29 18:00:56'),
(155, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-29 18:20:37'),
(156, 2, 'UPDATE DATA DIRI', 'User 000001 mengupdate data diri', '2025-10-29 18:21:28'),
(157, 2, 'UPDATE FOTO', 'User mengupdate foto profil', '2025-10-29 18:23:08'),
(158, 2, 'LOGOUT', 'User 000001 logout', '2025-10-29 18:23:31'),
(159, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-29 18:23:36'),
(160, 1, 'EXPORT DATA', 'Admin export data personel ke Excel', '2025-10-29 18:25:21'),
(161, 1, 'UPDATE DATA', 'Admin mengupdate data personel NRP: 000002', '2025-10-29 18:26:12'),
(162, 1, 'LOGOUT', 'User 123456 logout', '2025-10-29 18:30:26'),
(163, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-29 18:30:37'),
(164, 1, 'LOGOUT', 'User 123456 logout', '2025-10-29 18:43:35'),
(165, 2, 'LOGIN', 'User 000001 berhasil login', '2025-10-29 19:00:28'),
(166, 2, 'LOGOUT', 'User 000001 logout', '2025-10-29 19:00:35'),
(167, 1, 'LOGIN', 'User 123456 berhasil login', '2025-10-29 19:00:40'),
(168, 1, 'EXPORT DATA', 'Admin export data personel ke Excel', '2025-10-29 19:00:50'),
(169, 1, 'LOGOUT', 'User 123456 logout', '2025-10-29 19:01:29'),
(170, 2, 'LOGIN', 'User 000001 berhasil login', '2025-11-09 04:42:17'),
(171, 2, 'LOGOUT', 'User 000001 logout', '2025-11-09 04:42:25'),
(172, 1, 'LOGIN', 'User 123456 berhasil login', '2025-11-09 04:42:35'),
(173, 1, 'TAMBAH USER', 'Admin menambahkan user baru dengan NRP: 000004', '2025-11-09 04:46:20'),
(174, 1, 'LOGOUT', 'User 123456 logout', '2025-11-09 04:46:29'),
(177, 2, 'LOGIN', 'User 000001 berhasil login', '2025-11-09 04:48:13'),
(178, 2, 'LOGOUT', 'User 000001 logout', '2025-11-09 04:55:57'),
(179, 1, 'LOGIN', 'User 123456 berhasil login', '2025-11-09 04:56:02'),
(180, 1, 'LOGOUT', 'User 123456 logout', '2025-11-09 05:00:33'),
(183, 2, 'LOGIN', 'User 000001 berhasil login', '2025-11-09 05:01:05'),
(184, 2, 'LOGOUT', 'User 000001 logout', '2025-11-09 05:45:42'),
(185, 1, 'LOGIN', 'User 123456 berhasil login', '2025-11-09 05:45:47'),
(186, 1, 'EXPORT DATA', 'Admin export data personel ke Excel', '2025-11-09 05:47:18'),
(187, 1, 'LOGOUT', 'User 123456 logout', '2025-11-09 05:47:38'),
(190, 1, 'LOGIN', 'User 123456 berhasil login', '2025-11-09 05:48:21'),
(191, 1, 'LOGOUT', 'User 123456 logout', '2025-11-09 05:48:39'),
(192, 2, 'LOGIN', 'User 000001 berhasil login', '2025-11-09 05:48:58'),
(193, 2, 'LOGOUT', 'User 000001 logout', '2025-11-09 05:49:02'),
(194, 1, 'LOGIN', 'User 123456 berhasil login', '2025-11-09 05:49:06'),
(195, 1, 'EXPORT DATA', 'Admin export data personel ke Excel', '2025-11-09 05:49:20'),
(196, 1, 'LOGOUT', 'User 123456 logout', '2025-11-09 05:49:42'),
(199, 2, 'LOGIN', 'User 000001 berhasil login', '2025-11-23 17:44:42'),
(200, 2, 'LOGOUT', 'User 000001 logout', '2025-11-23 17:44:50'),
(201, 1, 'LOGIN', 'User 123456 berhasil login', '2025-11-23 17:44:56'),
(202, 1, 'LOGOUT', 'User 123456 logout', '2025-11-23 17:45:29'),
(203, 1, 'LOGIN', 'User 123456 berhasil login', '2025-11-23 17:45:37'),
(204, 1, 'LOGOUT', 'User 123456 logout', '2025-11-23 17:46:05'),
(205, 1, 'LOGIN', 'User 123456 berhasil login', '2025-11-23 17:46:12'),
(206, 1, 'LOGOUT', 'User 123456 logout', '2025-11-23 17:46:44'),
(207, 2, 'LOGIN', 'User 000001 berhasil login', '2025-11-23 17:46:50'),
(208, 2, 'LOGOUT', 'User 000001 logout', '2025-11-23 17:47:07'),
(213, 1, 'LOGIN', 'User 123456 berhasil login', '2025-11-23 17:49:10'),
(214, 1, 'EXPORT DATA', 'Admin export data personel ke Excel', '2025-11-23 18:10:05'),
(215, 1, 'IMPORT DATA', 'Admin mengimport 0 data personel dan membuat 0 user baru.', '2025-11-23 18:19:27'),
(216, 1, 'IMPORT DATA', 'Admin import excel: 0 sukses, 0 gagal.', '2025-11-23 18:22:48'),
(217, 1, 'IMPORT DATA', 'Admin import excel: 1 sukses, 0 gagal.', '2025-11-23 18:24:59'),
(218, 1, 'HAPUS DATA', 'Admin menghapus data personel NRP: 111111', '2025-11-23 18:25:08'),
(219, 1, 'IMPORT DATA', 'Admin import excel: 1 sukses, 0 gagal.', '2025-11-23 18:25:33'),
(220, 1, 'UPDATE DATA', 'Admin mengupdate data personel NRP: 000002', '2025-11-23 18:26:17'),
(221, 1, 'UPDATE DATA', 'Admin mengupdate data personel NRP: 000003', '2025-11-23 18:26:49'),
(222, 1, 'HAPUS DATA', 'Admin menghapus data personel NRP: 111111', '2025-11-23 18:26:58'),
(223, 1, 'TAMBAH USER', 'Admin menambahkan user baru dengan NRP: 000004', '2025-11-23 18:28:22'),
(224, 1, 'LOGOUT', 'User 123456 logout', '2025-11-23 18:28:37'),
(227, 2, 'LOGIN', 'User 000001 berhasil login', '2025-11-23 18:45:07'),
(228, 2, 'LOGOUT', 'User 000001 logout', '2025-11-23 18:57:37');

-- --------------------------------------------------------

--
-- Struktur dari tabel `korp`
--

CREATE TABLE `korp` (
  `KORPSID` varchar(2) NOT NULL,
  `SEBUTAN` varchar(3) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

--
-- Dumping data untuk tabel `korp`
--

INSERT INTO `korp` (`KORPSID`, `SEBUTAN`) VALUES
('12', 'P'),
('13', 'Pnb'),
('22', 'T'),
('23', 'Nav'),
('32', 'E'),
('33', 'Tek'),
('42', 'S'),
('43', 'Lek'),
('52', 'MAR'),
('53', 'Kal'),
('62', 'K'),
('63', 'Adm'),
('72', 'KH'),
('73', 'Kes'),
('82', 'PM'),
('83', 'Pas'),
('93', 'Sus'),
('A1', 'Inf'),
('A3', 'Pom'),
('B1', 'Kav'),
('C1', 'Arm'),
('D1', 'Arh'),
('E1', 'Czi'),
('F1', 'Cpm'),
('G1', 'Cba'),
('H1', 'CHK'),
('K1', 'Ckm'),
('M1', 'Cpl'),
('N1', 'Chb'),
('P1', 'Caj'),
('Q1', 'Cku'),
('R1', 'Ctp'),
('X1', 'TIT'),
('Y1', 'CPN'),
('Z1', 'TNI');

-- --------------------------------------------------------

--
-- Struktur dari tabel `matra`
--

CREATE TABLE `matra` (
  `MTR` int NOT NULL,
  `Nama` varchar(50) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

--
-- Dumping data untuk tabel `matra`
--

INSERT INTO `matra` (`MTR`, `Nama`) VALUES
(0, 'PNS'),
(1, 'TNI AD'),
(2, 'TNI AL'),
(3, 'TNI AU');

-- --------------------------------------------------------

--
-- Struktur dari tabel `pangkat`
--

CREATE TABLE `pangkat` (
  `kd_pkt` varchar(5) NOT NULL,
  `sebutan` varchar(50) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

--
-- Dumping data untuk tabel `pangkat`
--

INSERT INTO `pangkat` (`kd_pkt`, `sebutan`) VALUES
('11', 'I/A'),
('12', 'I/B'),
('13', 'I/C'),
('14', 'I/D'),
('21', 'II/A'),
('22', 'II/B'),
('23', 'II/C'),
('24', 'II/D'),
('31', 'III/A'),
('32', 'III/B'),
('33', 'III/C'),
('34', 'III/D'),
('41', 'IV/A'),
('42', 'IV/B'),
('43', 'IV/C'),
('44', 'IV/D'),
('45', 'IV/E'),
('51', 'Prada'),
('52', 'Pratu'),
('53', 'Praka'),
('54', 'Kopda'),
('55', 'Koptu'),
('56', 'Kopka'),
('61', 'Serda'),
('62', 'Sertu'),
('63', 'Serka'),
('64', 'Serma'),
('65', 'Pelda'),
('66', 'Peltu'),
('67', 'CAPA'),
('71', 'Letda'),
('72', 'Lettu'),
('73', 'Kapten'),
('81', 'Mayor'),
('82', 'Letkol'),
('83', 'Kolonel'),
('91', 'Brigjen TNI'),
('92', 'Mayjen TNI'),
('93', 'Letjen TNI'),
('94', 'Jenderal TNI');

-- --------------------------------------------------------

--
-- Struktur dari tabel `personel`
--

CREATE TABLE `personel` (
  `id` int NOT NULL,
  `nrp` char(6) NOT NULL,
  `nama` varchar(100) NOT NULL,
  `tempat_lahir` varchar(50) NOT NULL,
  `tanggal_lahir` date NOT NULL,
  `korp` varchar(2) DEFAULT NULL,
  `pangkat` varchar(5) DEFAULT NULL,
  `matra` int DEFAULT NULL,
  `kd_satker` varchar(5) DEFAULT NULL,
  `alamat` text,
  `nik` char(16) DEFAULT NULL,
  `no_hp` varchar(15) DEFAULT NULL,
  `no_kep` varchar(20) DEFAULT NULL,
  `no_sprint` varchar(20) DEFAULT NULL,
  `kd_gender` char(1) DEFAULT NULL,
  `satker_lama` varchar(10) DEFAULT NULL,
  `no_kep_lama` varchar(20) DEFAULT NULL,
  `no_sprint_lama` varchar(20) DEFAULT NULL,
  `foto_profil` varchar(255) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

--
-- Dumping data untuk tabel `personel`
--

INSERT INTO `personel` (`id`, `nrp`, `nama`, `tempat_lahir`, `tanggal_lahir`, `korp`, `pangkat`, `matra`, `kd_satker`, `alamat`, `nik`, `no_hp`, `no_kep`, `no_sprint`, `kd_gender`, `satker_lama`, `no_kep_lama`, `no_sprint_lama`, `foto_profil`) VALUES
(1, '000001', 'ARTHURITO LIANDO SIMANJUNTAK', 'Jakarta', '1998-05-03', 'C1', '73', 1, 'D13', 'Jl Dahlan Raya No 46, Harjamukti, Cimanggis, Depok', '3175051410031001', '089682564616', NULL, NULL, 'L', 'B02', NULL, NULL, '000001_1761762188.jpg'),
(3, '000002', 'RYAN DWI KURNIAWAN', 'Bogor', '1999-10-27', 'E1', '45', 1, 'D13', 'jl cileungsi', '3219011237183012', '081221714167', NULL, NULL, 'L', 'D05', NULL, NULL, '000002_1759724695.jpg'),
(4, '000003', 'ADITYA PRATAMA', 'Padang', '2000-02-20', '62', '73', 2, 'D13', 'Jl raya AL Jatimakmur', '3275012312312312', '085778606564', NULL, NULL, 'L', 'C06', NULL, NULL, '000003_1759724288.jpg');

-- --------------------------------------------------------

--
-- Struktur dari tabel `satker`
--

CREATE TABLE `satker` (
  `kd_satker` varchar(5) NOT NULL,
  `nama_satker` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

--
-- Dumping data untuk tabel `satker`
--

INSERT INTO `satker` (`kd_satker`, `nama_satker`) VALUES
('A00', 'SMIN PANGLIMA TNI'),
('A01', 'SMIN WAPANG TNI'),
('B01', 'SMIN KASUNG TNI'),
('B02', 'ITJEN TNI'),
('B03', 'STAF AHLI TNI'),
('B04', 'SRENUM TNI'),
('B05', 'SINTEL TNI'),
('B06', 'SOPS TNI'),
('B07', 'SPERS TNI'),
('B08', 'SLOG TNI'),
('B09', 'STER TNI'),
('B10', 'SKOMLEK TNI'),
('B11', 'PUSPOM TNI'),
('B12', 'PUSMINPERS TNI'),
('C01', 'SATKOMLEK TNI'),
('C02', 'PUSDALOPS TNI'),
('C03', 'SETUM TNI'),
('C04', 'DENMA MABES TNI'),
('C05', 'PUSJASPERMILDAS'),
('C06', 'KOOPSSUS TNI'),
('C07', 'PUS RB TNI'),
('C08', 'PUSPI TNI'),
('D01', 'SESKO TNI'),
('D02', 'KODIKLAT TNI'),
('D03', 'AKADEMI TNI'),
('D04', 'BAIS TNI'),
('D05', 'PASPAMPRES'),
('D06', 'BABINKUM TNI'),
('D07', 'PUSPEN TNI'),
('D08', 'PUSKES TNI'),
('D10', 'PUSBINTAL TNI'),
('D11', 'PUSKU TNI'),
('D12', 'PUSJARAH TNI'),
('D13', 'PUSINFOLAHTA TNI'),
('D14', 'PMPP TNI'),
('D15', 'PUSJIANGSTRALITBANG TNI'),
('D16', 'PUSKERSIN TNI'),
('D17', 'SATSIBER TNI'),
('D18', 'PUSINFOMAR'),
('D19', 'KOGARTAP I/JKT'),
('D20', 'KOGARTAP II/BDG'),
('D21', 'KOGARTAP III/SBY'),
('D22', 'KOGABWILHAN I'),
('D23', 'KOGABWILHAN II'),
('D24', 'KOGABWILHAN III'),
('D25', 'PUSADA TNI'),
('D26', 'BALOG TNI');

-- --------------------------------------------------------

--
-- Struktur dari tabel `users`
--

CREATE TABLE `users` (
  `id` int NOT NULL,
  `nrp` char(6) DEFAULT NULL,
  `password` varchar(255) NOT NULL,
  `role` enum('admin','user') DEFAULT 'user'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

--
-- Dumping data untuk tabel `users`
--

INSERT INTO `users` (`id`, `nrp`, `password`, `role`) VALUES
(1, '123456', '$2y$10$anPDSC5ximN0LYplAu6FquHAQFUzVJG.89.pI.J1qUvicbweiCwda', 'admin'),
(2, '000001', '$2y$10$anPDSC5ximN0LYplAu6FquHAQFUzVJG.89.pI.J1qUvicbweiCwda', 'user'),
(3, '000002', '$2y$10$anPDSC5ximN0LYplAu6FquHAQFUzVJG.89.pI.J1qUvicbweiCwda', 'user'),
(4, '000003', '$2y$10$/69MTLJzdmDgFMHj8aEW2eDPr8eya20c64vkldjM2s6nmK0NPVT.u', 'user');

--
-- Indexes for dumped tables
--

--
-- Indeks untuk tabel `gender`
--
ALTER TABLE `gender`
  ADD PRIMARY KEY (`kd_gender`);

--
-- Indeks untuk tabel `history_log`
--
ALTER TABLE `history_log`
  ADD PRIMARY KEY (`id`),
  ADD KEY `user_id` (`user_id`);

--
-- Indeks untuk tabel `korp`
--
ALTER TABLE `korp`
  ADD PRIMARY KEY (`KORPSID`);

--
-- Indeks untuk tabel `matra`
--
ALTER TABLE `matra`
  ADD PRIMARY KEY (`MTR`);

--
-- Indeks untuk tabel `pangkat`
--
ALTER TABLE `pangkat`
  ADD PRIMARY KEY (`kd_pkt`);

--
-- Indeks untuk tabel `personel`
--
ALTER TABLE `personel`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `nrp` (`nrp`),
  ADD KEY `korp` (`korp`),
  ADD KEY `pangkat` (`pangkat`),
  ADD KEY `matra` (`matra`),
  ADD KEY `satker` (`kd_satker`),
  ADD KEY `fk_personel_gender` (`kd_gender`);

--
-- Indeks untuk tabel `satker`
--
ALTER TABLE `satker`
  ADD PRIMARY KEY (`kd_satker`);

--
-- Indeks untuk tabel `users`
--
ALTER TABLE `users`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `nrp` (`nrp`);

--
-- AUTO_INCREMENT untuk tabel yang dibuang
--

--
-- AUTO_INCREMENT untuk tabel `history_log`
--
ALTER TABLE `history_log`
  MODIFY `id` int NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=229;

--
-- AUTO_INCREMENT untuk tabel `personel`
--
ALTER TABLE `personel`
  MODIFY `id` int NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=8;

--
-- AUTO_INCREMENT untuk tabel `users`
--
ALTER TABLE `users`
  MODIFY `id` int NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=9;

--
-- Ketidakleluasaan untuk tabel pelimpahan (Dumped Tables)
--

--
-- Ketidakleluasaan untuk tabel `history_log`
--
ALTER TABLE `history_log`
  ADD CONSTRAINT `history_log_ibfk_1` FOREIGN KEY (`user_id`) REFERENCES `users` (`id`);

--
-- Ketidakleluasaan untuk tabel `personel`
--
ALTER TABLE `personel`
  ADD CONSTRAINT `fk_personel_gender` FOREIGN KEY (`kd_gender`) REFERENCES `gender` (`kd_gender`),
  ADD CONSTRAINT `personel_ibfk_1` FOREIGN KEY (`nrp`) REFERENCES `users` (`nrp`),
  ADD CONSTRAINT `personel_ibfk_2` FOREIGN KEY (`korp`) REFERENCES `korp` (`KORPSID`),
  ADD CONSTRAINT `personel_ibfk_3` FOREIGN KEY (`pangkat`) REFERENCES `pangkat` (`kd_pkt`),
  ADD CONSTRAINT `personel_ibfk_4` FOREIGN KEY (`matra`) REFERENCES `matra` (`MTR`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
