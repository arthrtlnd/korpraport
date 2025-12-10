<?php
session_start();
require_once '../app/koneksi.php';
check_admin();

$success = '';
$error = '';

// --- Cek Pesan Session ---
if (isset($_SESSION['success'])) {
    $success = $_SESSION['success'];
    unset($_SESSION['success']);
}
if (isset($_SESSION['error'])) {
    $error = $_SESSION['error'];
    unset($_SESSION['error']);
}

// --- Handle Import Excel ---
if (isset($_POST['import_excel']) && isset($_FILES['excel_file'])) {
    if (!extension_loaded('zip')) {
        $error = "Gagal Import: Ekstensi PHP 'zip' belum aktif.";
    } else {
        require_once '../vendor/autoload.php'; 
        
        $file = $_FILES['excel_file'];
        $allowed_ext = ['xlsx', 'xls'];
        $ext = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
        
        if (in_array($ext, $allowed_ext)) {
            try {
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file['tmp_name']);
                $worksheet = $spreadsheet->getActiveSheet();
                $rows = $worksheet->toArray();
                
                $imported = 0;
                $failed = 0;
                $users_created = 0;
                $error_details = ""; 
                
                $stmt_check = mysqli_prepare($conn, "SELECT id FROM personel WHERE nrp = ?");
                $stmt_update = mysqli_prepare($conn, "UPDATE personel SET nama=?, nik=?, tempat_lahir=?, tanggal_lahir=?, pangkat=?, korp=?, matra=?, kd_satker=?, satker_lama=? WHERE id=?");
                $stmt_insert = mysqli_prepare($conn, "INSERT INTO personel (nrp, nama, nik, tempat_lahir, tanggal_lahir, pangkat, korp, matra, kd_satker, satker_lama) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
                
                $stmt_check_user = mysqli_prepare($conn, "SELECT id FROM users WHERE nrp = ?");
                $stmt_insert_user = mysqli_prepare($conn, "INSERT INTO users (nrp, password, role) VALUES (?, ?, ?)");

                for ($i = 1; $i < count($rows); $i++) {
                    $row = $rows[$i];
                    if (empty($row[0]) || empty($row[1])) continue; 

                    $nrp = clean_input($row[0]);
                    $nama = clean_input($row[1]);
                    $nik = clean_input($row[2] ?? null);
                    $tempat_lahir = clean_input($row[3] ?? '');
                    
                    $tanggal_lahir = null;
                    if (!empty($row[4])) {
                        if (is_numeric($row[4])) {
                            $tanggal_lahir = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($row[4])->format('Y-m-d');
                        } else {
                            $timestamp = strtotime($row[4]);
                            if ($timestamp) $tanggal_lahir = date('Y-m-d', $timestamp);
                        }
                    }

                    $pangkat = clean_input($row[5] ?? null);
                    $korp = clean_input($row[6] ?? null);
                    $matra = clean_input($row[7] ?? null);
                    $kd_satker = clean_input($row[8] ?? null);
                    $satker_lama = clean_input($row[9] ?? null);

                    // Buat User Login jika belum ada
                    mysqli_stmt_bind_param($stmt_check_user, "s", $nrp);
                    mysqli_stmt_execute($stmt_check_user);
                    $res_user_check = mysqli_stmt_get_result($stmt_check_user);

                    if (mysqli_num_rows($res_user_check) == 0) {
                        $default_password = password_hash('password', PASSWORD_DEFAULT);
                        $default_role = 'user';
                        mysqli_stmt_bind_param($stmt_insert_user, "sss", $nrp, $default_password, $default_role);
                        if (mysqli_stmt_execute($stmt_insert_user)) {
                            $users_created++;
                        } else {
                            $failed++;
                            $error_details .= "Baris " . ($i+1) . " (NRP $nrp): Gagal buat User. <br>";
                            continue;
                        }
                    }

                    // Proses Data Personel
                    mysqli_stmt_bind_param($stmt_check, "s", $nrp);
                    mysqli_stmt_execute($stmt_check);
                    $res_check = mysqli_stmt_get_result($stmt_check);
                    $existing = mysqli_fetch_assoc($res_check);

                    $result_query = false;
                    if ($existing) {
                        $pid = $existing['id'];
                        mysqli_stmt_bind_param($stmt_update, "sssssssssi", $nama, $nik, $tempat_lahir, $tanggal_lahir, $pangkat, $korp, $matra, $kd_satker, $satker_lama, $pid);
                        $result_query = mysqli_stmt_execute($stmt_update);
                    } else {
                        mysqli_stmt_bind_param($stmt_insert, "ssssssssss", $nrp, $nama, $nik, $tempat_lahir, $tanggal_lahir, $pangkat, $korp, $matra, $kd_satker, $satker_lama);
                        $result_query = mysqli_stmt_execute($stmt_insert);
                    }
                    
                    if ($result_query) $imported++;
                    else {
                        $failed++;
                        $error_details .= "Baris " . ($i+1) . " (NRP $nrp): " . mysqli_error($conn) . "<br>";
                    }
                }
                
                $success = "Import Selesai: $imported sukses, $failed gagal.";
                if ($users_created > 0) $success .= " ($users_created user baru dibuat).";
                if ($failed > 0) $error = "Detail Gagal:<br>" . $error_details;
                
                catat_log($_SESSION['user_id'], 'IMPORT DATA', "Import Excel: $imported sukses.");
            } catch (Exception $e) {
                $error = "Gagal import: " . $e->getMessage();
            }
        } else {
            $error = "Format file harus .xlsx atau .xls";
        }
    }
}

// --- FILTER & PAGINATION ---
$where_clause = "WHERE 1=1";
$join_clause = "";
$search = '';
$date_from = '';
$date_to = '';

// REVISI: Limit Options diubah jadi 5, 25, 50
$limit_options = [5, 25, 50];
$limit = 5; // Default

if (isset($_GET['limit']) && in_array($_GET['limit'], $limit_options)) $limit = (int)$_GET['limit'];

if (isset($_GET['search']) && !empty($_GET['search'])) {
    $search = clean_input($_GET['search']);
    $search_sql = escape_string($search);
    $where_clause .= " AND (p.nrp LIKE '%$search_sql%' OR p.nama LIKE '%$search_sql%' OR pkt.sebutan LIKE '%$search_sql%')";
}
if (isset($_GET['date_from']) && !empty($_GET['date_from'])) {
    $date_from = clean_input($_GET['date_from']);
    $where_clause .= " AND h_log.login_time >= '" . escape_string($date_from) . " 00:00:00'";
}
if (isset($_GET['date_to']) && !empty($_GET['date_to'])) {
    $date_to = clean_input($_GET['date_to']);
    $where_clause .= " AND h_log.login_time <= '" . escape_string($date_to) . " 23:59:59'";
}
if (!empty($date_from) || !empty($date_to)) {
    $join_clause = " INNER JOIN ( SELECT u.nrp, MAX(h.waktu) as login_time FROM history_log h JOIN users u ON h.user_id = u.id WHERE h.aksi = 'LOGIN' GROUP BY u.nrp ) AS h_log ON p.nrp = h_log.nrp ";
}

$filter_params = http_build_query(['search' => $search, 'date_from' => $date_from, 'date_to' => $date_to]);

$query_total = "SELECT COUNT(DISTINCT p.id) as total FROM personel p LEFT JOIN users u ON p.nrp = u.nrp LEFT JOIN pangkat pkt ON p.pangkat = pkt.kd_pkt $join_clause $where_clause";
$result_total = mysqli_query($conn, $query_total);
$total_personel_found = mysqli_fetch_assoc($result_total)['total'];

$query = "SELECT p.*, u.role, g.gender, k.SEBUTAN as korp_sebutan, pkt.sebutan as pangkat_sebutan, m.Nama as matra_nama, s.nama_satker, s_lama.nama_satker as nama_satker_lama FROM personel p LEFT JOIN users u ON p.nrp = u.nrp LEFT JOIN gender g ON p.kd_gender = g.kd_gender LEFT JOIN korp k ON p.korp = k.KORPSID LEFT JOIN pangkat pkt ON p.pangkat = pkt.kd_pkt LEFT JOIN matra m ON p.matra = m.MTR LEFT JOIN satker s ON p.kd_satker = s.kd_satker LEFT JOIN satker s_lama ON p.satker_lama = s_lama.kd_satker $join_clause $where_clause GROUP BY p.id ORDER BY p.id DESC LIMIT $limit";
$result = mysqli_query($conn, $query);
$total_personel_showing = mysqli_num_rows($result);
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Master Data Personel - KORPRAPORT</title>
    <link rel="stylesheet" href="../assets/css/masterpersonel.css">
    <style>
        .btn-download { background: #28a745; color: white; padding: 6px 12px; font-size: 12px; border: none; border-radius: 4px; cursor: pointer; text-decoration: none; transition: all 0.3s; display: inline-block; margin-top: 4px; }
        .btn-download:hover { background: #218838; }
        .alert-error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; padding: 15px; border-radius: 5px; margin-bottom: 20px; font-size: 14px; }
    </style>
</head>
<body>
    <div class="sidebar">
        <div class="sidebar-header">
            <h2>KORPRAPORT</h2>
            <p>Admin Panel</p>
        </div>
        <ul class="sidebar-menu">
            <li><a href="dashboard.php"><span>📊</span> Dashboard</a></li>
            <li><a href="masterpersonel.php" class="active"><span>👥</span> Master Data Personel</a></li>
            <li><a href="adduser.php"><span>➕</span> Tambah User</a></li>
            <li><a href="historylog.php"><span>📋</span> History Log</a></li>
            <li><a href="../auth/logout.php"><span>🚪</span> Logout</a></li>
        </ul>
    </div>
    
    <div class="main-content">
        <div class="top-bar">
            <h1>Master Data Personel</h1>
            <div class="user-info">
                <span>Menampilkan <strong><?php echo $total_personel_showing; ?></strong> dari <strong><?php echo $total_personel_found; ?></strong> data ditemukan</span>
            </div>
        </div>
        
        <?php if ($success): ?>
            <div class="alert alert-success"><?php echo $success; ?></div>
        <?php endif; ?>
        
        <?php if ($error): ?>
            <div class="alert alert-error"><?php echo $error; ?></div>
        <?php endif; ?>
        
        <div class="card">
            <div class="card-header">
                <div class="header-actions">
                    <button class="btn btn-success" onclick="document.getElementById('importModal').style.display='block'">
                        📥 Import Excel
                    </button>
                    <a href="export_excel.php?<?php echo $filter_params; ?>" class="btn btn-primary">📤 Export (Excel)</a>
                    <a href="adduser.php" class="btn btn-secondary">➕ Tambah User</a>
                </div>
            </div>
            
            <div class="filter-section">
                <form method="GET" action="" class="filter-form">
                    <div class="filter-group">
                        <label>Cari (NRP, Nama, Pangkat)</label>
                        <input type="text" name="search" placeholder="Cari..." value="<?php echo htmlspecialchars($search); ?>">
                    </div>
                    <div class="filter-group">
                        <label>Login Dari Tanggal:</label>
                        <input type="date" name="date_from" value="<?php echo htmlspecialchars($date_from); ?>">
                    </div>
                    <div class="filter-group">
                        <label>Login Sampai Tanggal:</label>
                        <input type="date" name="date_to" value="<?php echo htmlspecialchars($date_to); ?>">
                    </div>
                    <div class="filter-group">
                        <label>Tampilkan:</label>
                        <select name="limit" onchange="this.form.submit()">
                            <!-- REVISI: Opsi limit diperbarui -->
                            <option value="5" <?php echo $limit == 5 ? 'selected' : ''; ?>>5</option>
                            <option value="25" <?php echo $limit == 25 ? 'selected' : ''; ?>>25</option>
                            <option value="50" <?php echo $limit == 50 ? 'selected' : ''; ?>>50</option>
                        </select>
                    </div>
                    <button type="submit" class="btn btn-search">🔍 Filter</button>
                    <a href="masterpersonel.php" class="btn btn-reset">🔄 Reset</a>
                </form>
            </div>
            
            <div class="table-container">
                <table id="personelTable">
                    <thead>
                        <tr>
                            <th>No</th>
                            <th>Nama</th>
                            <th>NRP</th>
                            <th>Pangkat</th>
                            <th>Korp</th>
                            <th>Matra</th>
                            <th>Jenis Kelamin</th>
                            <th>Tempat Lahir</th>
                            <th>Tanggal Lahir</th>
                            <th>Satuan Lama</th>
                            <th>No KEP Lama</th>
                            <th>No Sprint Lama</th>
                            <th>Satuan Baru</th>
                            <th>No KEP Baru</th>
                            <th>No Sprint Baru</th>
                            <th>NIK</th>
                            <th>Alamat</th>
                            <th>No HP</th>
                            <th>Role</th>
                            <th>Aksi</th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php 
                        $no = 1;
                        while ($row = mysqli_fetch_assoc($result)): 
                        ?>
                        <tr>
                            <td><?php echo $no++; ?></td>
                            <td><?php echo htmlspecialchars($row['nama']); ?></td>
                            <td><?php echo htmlspecialchars($row['nrp']); ?></td>
                            <td><?php echo htmlspecialchars($row['pangkat_sebutan'] ?? '-'); ?></td>
                            <td><?php echo htmlspecialchars($row['korp_sebutan'] ?? '-'); ?></td>
                            <td><?php echo htmlspecialchars($row['matra_nama'] ?? '-'); ?></td>
                            <td><?php echo htmlspecialchars($row['gender'] ?? '-'); ?></td>
                            <td><?php echo htmlspecialchars($row['tempat_lahir'] ?? '-'); ?></td>
                            <td><?php echo $row['tanggal_lahir'] ? date('d/m/Y', strtotime($row['tanggal_lahir'])) : '-'; ?></td>
                            <td><?php echo htmlspecialchars($row['nama_satker_lama'] ?? '-'); ?></td>
                            <td><?php echo htmlspecialchars($row['no_kep_lama'] ?? '-'); ?></td>
                            <td><?php echo htmlspecialchars($row['no_sprint_lama'] ?? '-'); ?></td>
                            <td><?php echo htmlspecialchars($row['nama_satker'] ?? '-'); ?></td>
                            <td><?php echo htmlspecialchars($row['no_kep'] ?? '-'); ?></td>
                            <td><?php echo htmlspecialchars($row['no_sprint'] ?? '-'); ?></td>
                            <td><?php echo htmlspecialchars($row['nik'] ?? '-'); ?></td>
                            <td><?php echo htmlspecialchars($row['alamat'] ?? '-'); ?></td>
                            <td><?php echo htmlspecialchars($row['no_hp'] ?? '-'); ?></td>
                            <td>
                                <span class="badge badge-<?php echo strtolower($row['role'] ?? 'user'); ?>">
                                    <?php echo strtoupper($row['role'] ?? 'USER'); ?>
                                </span>
                            </td>
                            <td>
                                <div class="action-buttons">
                                    <a href="edituser.php?id=<?php echo $row['id']; ?>" class="btn-edit">Edit</a>
                                    <a href="deleteuser.php?id=<?php echo $row['id']; ?>" class="btn-delete" onclick="return confirm('Yakin ingin menghapus data ini?')">Hapus</a>
                                    <?php if (!empty($row['foto_profil'])): ?>
                                        <a href="../uploads/profile/<?php echo $row['foto_profil']; ?>" class="btn-download" download="FOTO_<?php echo $row['nrp']; ?>">Unduh Foto</a>
                                    <?php endif; ?>
                                </div>
                            </td>
                        </tr>
                        <?php endwhile; ?>
                        
                        <?php if ($total_personel_showing == 0): ?>
                        <tr>
                            <td colspan="20" style="text-align: center; padding: 30px; color: #999;">Tidak ada data yang ditemukan</td>
                        </tr>
                        <?php endif; ?>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <!-- Import Modal -->
    <div id="importModal" class="modal">
        <div class="modal-content">
            <span class="close" onclick="document.getElementById('importModal').style.display='none'">&times;</span>
            <h2>Import Data dari Excel</h2>
            <p>Pastikan format kolom sesuai dengan template.</p>
            <form method="POST" enctype="multipart/form-data">
                <div class="form-group">
                    <label>Pilih File Excel (.xlsx atau .xls)</label>
                    <input type="file" name="excel_file" accept=".xlsx,.xls" required>
                </div>
                <button type="submit" name="import_excel" class="btn btn-primary">Upload & Import</button>
            </form>
            <div style="margin-top: 15px;">
                <a href="download_template.php" class="btn btn-secondary">📥 Download Template Excel</a>
            </div>
        </div>
    </div>
    
    <script>
        window.onclick = function(event) {
            const modal = document.getElementById('importModal');
            if (event.target == modal) {
                modal.style.display = "none";
            }
        }
    </script>
</body>
</html>