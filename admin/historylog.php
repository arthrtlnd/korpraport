<?php
session_start();
require_once '../app/koneksi.php';
check_admin();

// REVISI 3: Tambah filter waktu
$where_clause = "WHERE 1=1";
$date_from = '';
$date_to = '';
$search = '';

if (isset($_GET['search']) && !empty($_GET['search'])) {
    $search = clean_input($_GET['search']);
    $where_clause .= " AND (u.nrp LIKE '%$search%' OR p.nama LIKE '%$search%' OR h.aksi LIKE '%$search%' OR h.keterangan LIKE '%$search%')";
}

if (isset($_GET['date_from']) && !empty($_GET['date_from'])) {
    $date_from = clean_input($_GET['date_from']);
    // Tambah jam 00:00:00
    $where_clause .= " AND h.waktu >= '$date_from 00:00:00'";
}

if (isset($_GET['date_to']) && !empty($_GET['date_to'])) {
    $date_to = clean_input($_GET['date_to']);
    // Tambah jam 23:59:59
    $where_clause .= " AND h.waktu <= '$date_to 23:59:59'";
}


// Query untuk mengambil history log
$query = "SELECT h.*, u.nrp, p.nama 
          FROM history_log h
          LEFT JOIN users u ON h.user_id = u.id
          LEFT JOIN personel p ON u.nrp = p.nrp
          $where_clause
          ORDER BY h.waktu DESC";

$result = mysqli_query($conn, $query);
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>History Log - KORPRAPORT</title>
    <!-- REVISI 4: Ganti CSS -->
    <link rel="stylesheet" href="../assets/css/masterpersonel.css"> <!-- Pakai CSS master -->
    <link rel="stylesheet" href="../assets/css/history.css"> <!-- CSS spesifik -->
    <style>
        /* Override beberapa style masterpersonel.css jika perlu */
        .main-content {
            padding-top: 30px; /* Sesuaikan padding */
        }
        .filter-form { /* Style untuk filter baru */
            display: flex;
            gap: 15px;
            align-items: flex-end;
            flex-wrap: wrap;
            background: #f8f9fa;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        .filter-group {
            display: flex;
            flex-direction: column;
            gap: 5px;
        }
        .filter-group label {
            font-size: 12px;
            color: #666;
            font-weight: 500;
        }
        .filter-group input {
            padding: 10px;
            border: 2px solid #ddd;
            border-radius: 5px;
            font-size: 14px;
        }
        /* Badge style tambahan */
        .badge-import {
            background: #007bff;
            color: white;
        }
    </style>
</head>
<body>
    <!-- REVISI 4: Ganti Navbar dengan Sidebar -->
    <div class="sidebar">
        <div class="sidebar-header">
            <h2>KORPRAPORT</h2>
            <p>Admin Panel</p>
        </div>
        <ul class="sidebar-menu">
            <li><a href="dashboard.php"><span>📊</span> Dashboard</a></li>
            <li><a href="masterpersonel.php"><span>👥</span> Master Data Personel</a></li>
            <li><a href="adduser.php"><span>➕</span> Tambah User</a></li>
            <li><a href="historylog.php" class="active"><span>📋</span> History Log</a></li>
            <li><a href="../auth/logout.php"><span>🚪</span> Logout</a></li>
        </ul>
    </div>
    
    <!-- REVISI 4: Bungkus konten dengan main-content -->
    <div class="main-content">
        <div class="top-bar">
            <h1>Riwayat Aktivitas Sistem</h1>
            <div class="user-info">
                 <a href="dashboard.php" class="btn btn-secondary">Kembali ke Dashboard</a>
            </div>
        </div>

        <div class="card">
            
            <!-- REVISI 3: Form Filter Waktu -->
            <form method="GET" action="" class="filter-form">
                <div class="filter-group">
                    <label>Cari (NRP, Aksi, Ket.)</label>
                    <input type="text" id="searchInput" name="search" placeholder="Cari..." value="<?php echo $search; ?>">
                </div>
                <div class="filter-group">
                    <label>Dari Tanggal:</label>
                    <input type="date" name="date_from" value="<?php echo $date_from; ?>">
                </div>
                <div class="filter-group">
                    <label>Sampai Tanggal:</label>
                    <input type="date" name="date_to" value="<?php echo $date_to; ?>">
                </div>
                <button type="submit" class="btn btn-search" style="height: 42px;">🔍 Filter</button>
                <a href="historylog.php" class="btn btn-reset" style="height: 42px;">🔄 Reset</a>
            </form>
            
            <div style="overflow-x: auto;">
                <table id="logTable">
                    <thead>
                        <tr>
                            <th>No</th>
                            <th>Waktu</th>
                            <th>NRP</th>
                            <th>Nama</th>
                            <th>Aksi</th>
                            <th>Keterangan</th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php 
                        $no = 1;
                        if (mysqli_num_rows($result) > 0):
                            while ($row = mysqli_fetch_assoc($result)): 
                                // Tentukan badge berdasarkan aksi
                                $badge_class = 'badge-tambah';
                                if (strpos(strtoupper($row['aksi']), 'LOGIN') !== false) {
                                    $badge_class = 'badge-login';
                                } elseif (strpos(strtoupper($row['aksi']), 'LOGOUT') !== false) {
                                    $badge_class = 'badge-logout';
                                } elseif (strpos(strtoupper($row['aksi']), 'UPDATE') !== false) {
                                    $badge_class = 'badge-update';
                                } elseif (strpos(strtoupper($row['aksi']), 'HAPUS') !== false) {
                                    $badge_class = 'badge-hapus';
                                } elseif (strpos(strtoupper($row['aksi']), 'IMPORT') !== false) {
                                    $badge_class = 'badge-import'; /* Tambah style jika perlu */
                                }
                        ?>
                        <tr>
                            <td><?php echo $no++; ?></td>
                            <td><?php echo date('d/m/Y H:i:s', strtotime($row['waktu'])); ?></td>
                            <td><?php echo $row['nrp'] ?? '-'; ?></td>
                            <td><?php echo $row['nama'] ?? 'Admin/Sistem'; ?></td>
                            <td>
                                <span class="badge <?php echo $badge_class; ?>">
                                    <?php echo $row['aksi']; ?>
                                </span>
                            </td>
                            <td><?php echo $row['keterangan']; ?></td>
                        </tr>
                        <?php 
                            endwhile;
                        else:
                        ?>
                        <tr>
                            <td colspan="6" class="no-data">
                                Belum ada riwayat aktivitas
                            </td>
                        </tr>
                        <?php endif; ?>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <!-- Hapus script lama, karena filter kini via PHP -->
    <!-- <script src="../assets/js/history.js"></script> -->
</body>
</html>

