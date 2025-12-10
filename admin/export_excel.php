<?php
session_start();
require_once '../app/koneksi.php';
check_admin();

// --- START REVISI 6: FILTER LOGIC ---
$where_clause = "WHERE 1=1";
$join_clause = "";
$search = '';
$date_from = '';
$date_to = '';

// Filter Search (NRP, Nama, Pangkat)
if (isset($_GET['search']) && !empty($_GET['search'])) {
    $search = clean_input($_GET['search']);
    $search_sql = escape_string($search);
    $where_clause .= " AND (p.nrp LIKE '%$search_sql%' OR p.nama LIKE '%$search_sql%' OR pkt.sebutan LIKE '%$search_sql%')";
}

// Filter Tanggal Login
if (isset($_GET['date_from']) && !empty($_GET['date_from'])) {
    $date_from = clean_input($_GET['date_from']);
    $where_clause .= " AND h_log.login_time >= '" . escape_string($date_from) . " 00:00:00'";
}

if (isset($_GET['date_to']) && !empty($_GET['date_to'])) {
    $date_to = clean_input($_GET['date_to']);
    $where_clause .= " AND h_log.login_time <= '" . escape_string($date_to) . " 23:59:59'";
}

// Jika filter tanggal login digunakan, tambahkan JOIN
if (!empty($date_from) || !empty($date_to)) {
    $join_clause = "
        INNER JOIN (
            SELECT u.nrp, MAX(h.waktu) as login_time
            FROM history_log h
            JOIN users u ON h.user_id = u.id
            WHERE h.aksi = 'LOGIN'
            GROUP BY u.nrp
        ) AS h_log ON p.nrp = h_log.nrp
    ";
}
// --- END REVISI 6 ---


// Query untuk mengambil semua data personel
// REVISI 6: Terapkan filter ke query
$query = "SELECT p.*, g.gender, k.SEBUTAN as korp_sebutan, 
          pkt.sebutan as pangkat_sebutan, m.Nama as matra_nama, 
          s.nama_satker, s_lama.nama_satker as nama_satker_lama
          FROM personel p
          LEFT JOIN gender g ON p.kd_gender = g.kd_gender
          LEFT JOIN korp k ON p.korp = k.KORPSID
          LEFT JOIN pangkat pkt ON p.pangkat = pkt.kd_pkt
          LEFT JOIN matra m ON p.matra = m.MTR
          LEFT JOIN satker s ON p.kd_satker = s.kd_satker
          LEFT JOIN satker s_lama ON p.satker_lama = s_lama.kd_satker
          $join_clause
          $where_clause
          GROUP BY p.id
          ORDER BY p.nama ASC"; // Urutkan berdasarkan nama

$result = mysqli_query($conn, $query);

// Set header untuk download Excel
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="Data_Personel_' . date('Y-m-d_His') . '.xls"');
header('Cache-Control: max-age=0');

?>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        table {
            border-collapse: collapse;
            width: 100%;
        }
        th, td {
            border: 1px solid black;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #7d1c1c;
            color: white;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <h2>DATA PERSONEL TNI</h2>
    <p>Tanggal Export: <?php echo date('d/m/Y H:i:s'); ?></p>
    
    <table>
        <thead>
            <tr>
                <th>No</th>
                <th>Nama</th>
                <th>NRP</th>
                <th>Pangkat</th>
                <th>Kode Pangkat</th> <!-- REVISI 7: Tambah Kolom -->
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
            </tr>
        </thead>
        <tbody>
            <?php 
            $no = 1;
            while ($row = mysqli_fetch_assoc($result)): 
            ?>
            <tr>
                <td><?php echo $no++; ?></td>
                <td><?php echo $row['nama']; ?></td>
                <td>'<?php echo $row['nrp']; ?></td> <!-- Tambah petik agar NRP dibaca sebagai teks -->
                <td><?php echo $row['pangkat_sebutan'] ?? '-'; ?></td>
                <td>'<?php echo $row['pangkat'] ?? '-'; ?></td> <!-- REVISI 7: Tambah Data -->
                <td><?php echo $row['korp_sebutan'] ?? '-'; ?></td>
                <td><?php echo $row['matra_nama'] ?? '-'; ?></td>
                <td><?php echo $row['gender'] ?? '-'; ?></td>
                <td><?php echo $row['tempat_lahir'] ?? '-'; ?></td>
                <td><?php echo $row['tanggal_lahir'] ? date('d/m/Y', strtotime($row['tanggal_lahir'])) : '-'; ?></td>
                <td><?php echo $row['nama_satker_lama'] ?? '-'; ?></td>
                <td>'<?php echo $row['no_kep_lama'] ?? '-'; ?></td>
                <td>'<?php echo $row['no_sprint_lama'] ?? '-'; ?></td>
                <td><?php echo $row['nama_satker'] ?? '-'; ?></td> <!-- Satuan Baru -->
                <td>'<?php echo $row['no_kep'] ?? '-'; ?></td> <!-- No KEP Baru -->
                <td>'<?php echo $row['no_sprint'] ?? '-'; ?></td> <!-- No Sprint Baru -->
                <td>'<?php echo $row['nik'] ?? '-'; ?></td>
                <td><?php echo $row['alamat'] ?? '-'; ?></td>
                <td>'<?php echo $row['no_hp'] ?? '-'; ?></td>
            </tr>
            <?php endwhile; ?>
        </tbody>
    </table>
</body>
</html>
<?php
catat_log($_SESSION['user_id'], 'EXPORT DATA', 'Admin export data personel ke Excel');
exit();
?>