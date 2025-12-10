<?php
session_start();
require_once '../app/koneksi.php';
check_admin();

// --- 1. SETUP DEFAULT TANGGAL (BULAN INI) ---
$start_date = isset($_GET['start_date']) ? clean_input($_GET['start_date']) : date('Y-m-01');
$end_date   = isset($_GET['end_date'])   ? clean_input($_GET['end_date'])   : date('Y-m-t');

// Tangkap parameter tanggal detail (jika user klik grafik)
$detail_date = isset($_GET['detail_date']) ? clean_input($_GET['detail_date']) : null;

// --- 2. QUERY GRAFIK PANGKAT (Bar Chart) ---
$query_pangkat = "SELECT pkt.sebutan, COUNT(p.id) as jumlah
                  FROM pangkat pkt
                  LEFT JOIN personel p ON p.pangkat = pkt.kd_pkt
                  GROUP BY pkt.kd_pkt, pkt.sebutan
                  ORDER BY pkt.kd_pkt DESC"; 

$result_pangkat = mysqli_query($conn, $query_pangkat);
$pangkat_labels = [];
$pangkat_data = [];
$total_personel = 0;

while ($row = mysqli_fetch_assoc($result_pangkat)) {
    $pangkat_labels[] = $row['sebutan'];
    $pangkat_data[] = (int)$row['jumlah'];
    $total_personel += (int)$row['jumlah'];
}

// --- 3. QUERY GRAFIK TREN LOGIN (Line Chart) ---
// REVISI PENTING: Menggunakan COUNT(DISTINCT user_id) 
// Agar yang dihitung adalah JUMLAH ORANGNYA, bukan jumlah aktivitas loginnya.
// Jadi kalau 1 orang login 10x, di grafik tetap dihitung 1.
$query_login = "SELECT DATE(waktu) as tanggal, COUNT(DISTINCT user_id) as jumlah 
                FROM history_log 
                WHERE aksi = 'LOGIN' 
                AND DATE(waktu) BETWEEN '$start_date' AND '$end_date'
                GROUP BY tanggal 
                ORDER BY tanggal ASC";
$result_login = mysqli_query($conn, $query_login);

$login_labels = [];
$login_data = [];
$login_dates_full = []; 

while ($row = mysqli_fetch_assoc($result_login)) {
    $date_obj = date_create($row['tanggal']);
    $login_labels[] = date_format($date_obj, "d M"); 
    $login_dates_full[] = $row['tanggal']; 
    $login_data[] = (int)$row['jumlah'];
}

// --- 4. QUERY DETAIL LOGIN (Tabel) ---
$where_detail = "WHERE h.aksi = 'LOGIN' ";
$table_title = "";

if ($detail_date) {
    // Mode Spesifik Tanggal (Hasil Klik Grafik)
    $where_detail .= " AND DATE(h.waktu) = '$detail_date'";
    $table_title = "Detail Login Tanggal: <span style='color:#7d1c1c;'>" . date('d F Y', strtotime($detail_date)) . "</span>";
} else {
    // Mode Default (Range Tanggal)
    $where_detail .= " AND DATE(h.waktu) BETWEEN '$start_date' AND '$end_date'";
    $table_title = "Daftar User yang Login (" . date('d M', strtotime($start_date)) . " - " . date('d M', strtotime($end_date)) . ")";
}

// Query ini sudah benar (GROUP BY nrp, tanggal) untuk menampilkan user unik per hari
$query_detail = "SELECT MAX(h.waktu) as waktu_terakhir, u.nrp, p.nama, pkt.sebutan as pangkat_sebutan
                 FROM history_log h
                 LEFT JOIN users u ON h.user_id = u.id
                 LEFT JOIN personel p ON u.nrp = p.nrp
                 LEFT JOIN pangkat pkt ON p.pangkat = pkt.kd_pkt
                 $where_detail
                 GROUP BY u.nrp, DATE(h.waktu) 
                 ORDER BY waktu_terakhir DESC";

$result_detail = mysqli_query($conn, $query_detail);
$total_rows = mysqli_num_rows($result_detail);
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Admin - KORPRAPORT</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', sans-serif; background: #f5f6fa; display: flex; min-height: 100vh; }
        
        /* Sidebar */
        .sidebar { width: 250px; background: linear-gradient(180deg, #7d1c1c 0%, #5d1717 100%); color: white; padding: 20px 0; position: fixed; height: 100vh; z-index: 100; }
        .sidebar-header { padding: 0 20px 20px; border-bottom: 1px solid rgba(255,255,255,0.1); margin-bottom: 20px; }
        .sidebar-menu { list-style: none; }
        .sidebar-menu a { display: block; padding: 12px 20px; color: white; text-decoration: none; border-left: 3px solid transparent; }
        .sidebar-menu a.active, .sidebar-menu a:hover { background: rgba(255,255,255,0.1); border-left-color: #FFB700; }
        .sidebar-menu span { margin-right: 10px; }

        /* Content Area */
        .main-content { margin-left: 250px; flex: 1; padding: 30px; width: calc(100% - 250px); }
        .top-bar { background: white; padding: 20px 30px; border-radius: 10px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 10px rgba(0,0,0,0.05); }
        
        /* Card Component */
        .card { background: white; padding: 25px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); margin-bottom: 25px; }
        .card h3 { margin-bottom: 20px; color: #333; font-size: 18px; border-bottom: 1px solid #eee; padding-bottom: 15px; display: flex; justify-content: space-between; align-items: center; }
        .card-subtitle { font-size: 14px; color: #666; font-weight: normal; }

        /* Filter Box */
        .filter-container { background: #fff; padding: 20px; border-radius: 10px; margin-bottom: 25px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 2px 10px rgba(0,0,0,0.05); border-left: 5px solid #FFB700; }
        .filter-form { display: flex; gap: 15px; align-items: flex-end; }
        .form-group { display: flex; flex-direction: column; gap: 5px; }
        .form-group label { font-size: 12px; font-weight: bold; color: #555; }
        .form-group input { padding: 8px 12px; border: 1px solid #ddd; border-radius: 5px; }
        .btn-filter { background: #7d1c1c; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-weight: bold; }
        .btn-filter:hover { background: #5d1717; }

        /* Table */
        .table-container { overflow-x: auto; max-height: 500px; }
        table { width: 100%; border-collapse: collapse; font-size: 14px; }
        th { text-align: left; padding: 12px; background: #f8f9fa; border-bottom: 2px solid #eee; color: #444; position: sticky; top: 0; }
        td { padding: 12px; border-bottom: 1px solid #eee; color: #555; }
        tr:hover { background: #fcfcfc; }
        
        .badge-admin { background: #2c3e50; color: white; padding: 3px 8px; border-radius: 4px; font-size: 11px; }
        .btn-reset-filter { background: #6c757d; color: white; border: none; padding: 5px 15px; border-radius: 5px; cursor: pointer; font-size: 12px; text-decoration: none; margin-left: 10px; }
        .btn-reset-filter:hover { background: #5a6268; }

        @media (max-width: 768px) { 
            .sidebar { transform: translateX(-250px); transition: 0.3s; }
            .main-content { margin-left: 0; width: 100%; }
            .filter-container { flex-direction: column; align-items: flex-start; gap: 15px; }
            .filter-form { flex-direction: column; width: 100%; }
            .form-group { width: 100%; }
        }
    </style>
</head>
<body>
    <div class="sidebar">
        <div class="sidebar-header">
            <h2>KORPRAPORT</h2>
            <p>Admin Panel</p>
        </div>
        <ul class="sidebar-menu">
            <li><a href="dashboard.php" class="active"><span>📊</span> Dashboard</a></li>
            <li><a href="masterpersonel.php"><span>👥</span> Master Data Personel</a></li>
            <li><a href="adduser.php"><span>➕</span> Tambah User</a></li>
            <li><a href="historylog.php"><span>📋</span> History Log</a></li>
            <li><a href="../auth/logout.php"><span>🚪</span> Logout</a></li>
        </ul>
    </div>
    
    <div class="main-content">
        <div class="top-bar">
            <h1>Dashboard Monitoring</h1>
            <div>Halo, <strong>Admin</strong></div>
        </div>

        <!-- 1. GRAFIK PANGKAT -->
        <div class="card">
            <h3>
                <div>📊 Distribusi Personel</div>
                <span class="card-subtitle">Total: <strong><?php echo $total_personel; ?></strong> Orang</span>
            </h3>
            <div style="height: 350px; position: relative;">
                <canvas id="pangkatChart"></canvas>
            </div>
        </div>

        <!-- 2. FILTER TANGGAL -->
        <div class="filter-container">
            <div>
                <strong>Filter Periode Laporan</strong>
                <p style="font-size: 12px; color: #666; margin-top: 5px;">Klik titik pada grafik login untuk melihat detail per hari.</p>
            </div>
            <form class="filter-form" method="GET">
                <div class="form-group">
                    <label>Dari Tanggal</label>
                    <input type="date" name="start_date" value="<?php echo $start_date; ?>" required>
                </div>
                <div class="form-group">
                    <label>Sampai Tanggal</label>
                    <input type="date" name="end_date" value="<?php echo $end_date; ?>" required>
                </div>
                <button type="submit" class="btn-filter">Tampilkan Data</button>
            </form>
        </div>

        <!-- 3. GRAFIK LOGIN (INTERAKTIF) -->
        <div class="card">
            <h3>
                📈 Grafik Aktivitas Login
                <span class="card-subtitle">Periode: <?php echo date('d M Y', strtotime($start_date)); ?> - <?php echo date('d M Y', strtotime($end_date)); ?></span>
            </h3>
            <div style="height: 300px; position: relative;">
                <canvas id="loginChart"></canvas>
            </div>
        </div>

        <!-- 4. TABEL DETAIL LOGIN -->
        <div class="card" id="detailSection">
            <h3>
                <div>
                    📋 <?php echo $table_title; ?>
                    <span class="card-subtitle" style="display:block; margin-top:5px; font-size:12px;">(Menampilkan login terakhir per hari untuk setiap user)</span>
                </div>
                <!-- Tombol Reset jika sedang mode Detail (Hasil Klik) -->
                <?php if ($detail_date): ?>
                    <a href="dashboard.php?start_date=<?php echo $start_date; ?>&end_date=<?php echo $end_date; ?>#detailSection" class="btn-reset-filter">🔄 Lihat Semua Tanggal</a>
                <?php endif; ?>
            </h3>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th width="5%">No</th>
                            <th width="20%">Waktu Login</th>
                            <th width="20%">NRP</th>
                            <th width="30%">Nama Personel</th>
                            <th width="25%">Pangkat / Keterangan</th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php 
                        if ($total_rows > 0): 
                            $no = 1;
                            while ($row = mysqli_fetch_assoc($result_detail)):
                                // Logika Nama Admin
                                $is_admin = ($row['nrp'] == '123456');
                                $display_nama = $is_admin ? 'Administrator' : ($row['nama'] ?? 'User Belum Lengkapi Data');
                                $display_pangkat = $is_admin ? '<span class="badge-admin">ADMIN SYSTEM</span>' : ($row['pangkat_sebutan'] ?? '-');
                        ?>
                        <tr>
                            <td><?php echo $no++; ?></td>
                            <td><?php echo date('d/m/Y H:i', strtotime($row['waktu_terakhir'])); ?> WIB</td>
                            <td><?php echo $row['nrp']; ?></td>
                            <td style="font-weight: 600;"><?php echo $display_nama; ?></td>
                            <td><?php echo $display_pangkat; ?></td>
                        </tr>
                        <?php endwhile; ?>
                        <?php else: ?>
                        <tr>
                            <td colspan="5" style="text-align:center; padding: 30px; color: #999;">
                                Tidak ada aktivitas login <?php echo $detail_date ? 'pada tanggal ini' : 'pada rentang tanggal ini'; ?>.
                            </td>
                        </tr>
                        <?php endif; ?>
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <script>
        // Data tanggal lengkap dari PHP untuk mapping klik
        const loginDatesFull = <?php echo json_encode($login_dates_full); ?>;
        
        // --- GRAFIK 1: PANGKAT (BAR) ---
        Chart.register(ChartDataLabels);
        const ctxPangkat = document.getElementById('pangkatChart').getContext('2d');
        new Chart(ctxPangkat, {
            type: 'bar',
            data: {
                labels: <?php echo json_encode($pangkat_labels); ?>,
                datasets: [{
                    label: 'Jumlah Personel',
                    data: <?php echo json_encode($pangkat_data); ?>,
                    backgroundColor: '#7d1c1c',
                    borderRadius: 4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    datalabels: {
                        anchor: 'end', align: 'top', color: '#333', font: { weight: 'bold' },
                        formatter: function(value) { return value > 0 ? value : ''; }
                    },
                    legend: { display: false }
                },
                scales: {
                    y: { beginAtZero: true, grid: { color: '#f0f0f0' } },
                    x: { grid: { display: false } }
                }
            }
        });

        // --- GRAFIK 2: LOGIN (LINE) - INTERAKTIF ---
        const ctxLogin = document.getElementById('loginChart').getContext('2d');
        new Chart(ctxLogin, {
            type: 'line',
            data: {
                labels: <?php echo json_encode($login_labels); ?>,
                datasets: [{
                    label: 'User Login', // Label diubah jadi User Login (unik)
                    data: <?php echo json_encode($login_data); ?>,
                    borderColor: '#FFB700',
                    backgroundColor: 'rgba(255, 183, 0, 0.1)',
                    borderWidth: 3,
                    pointBackgroundColor: '#7d1c1c',
                    pointRadius: 6, 
                    pointHoverRadius: 8,
                    fill: true,
                    tension: 0.3
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                // --- EVENT KLIK ---
                onClick: (e, activeElements, chart) => {
                    if (activeElements.length > 0) {
                        const index = activeElements[0].index;
                        const selectedDate = loginDatesFull[index];
                        
                        const currentUrl = new URL(window.location.href);
                        currentUrl.searchParams.set('detail_date', selectedDate);
                        
                        if (!currentUrl.searchParams.has('start_date')) {
                            currentUrl.searchParams.set('start_date', '<?php echo $start_date; ?>');
                            currentUrl.searchParams.set('end_date', '<?php echo $end_date; ?>');
                        }
                        
                        currentUrl.hash = "detailSection";
                        
                        window.location.href = currentUrl.toString();
                    }
                },
                onHover: (event, chartElement) => {
                    event.native.target.style.cursor = chartElement[0] ? 'pointer' : 'default';
                },
                plugins: {
                    datalabels: { display: false }, 
                    legend: { display: false },
                    tooltip: {
                        displayColors: false,
                        callbacks: {
                            label: function(context) {
                                return 'Klik untuk lihat ' + context.parsed.y + ' user';
                            }
                        }
                    }
                },
                scales: {
                    y: { 
                        beginAtZero: true, 
                        ticks: { stepSize: 1 }, 
                        grid: { color: '#f0f0f0' }
                    },
                    x: { grid: { display: false } }
                }
            }
        });
    </script>
</body>
</html>