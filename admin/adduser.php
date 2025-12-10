<?php
session_start();
require_once '../app/koneksi.php';
check_admin();

$success = '';
$error = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // SECURITY FIX: Validasi CSRF Token
    verify_csrf_token($_POST['csrf_token'] ?? '');

    $nrp = clean_input($_POST['nrp']);
    $password = $_POST['password'];
    $confirm_password = $_POST['confirm_password'];
    $role = clean_input($_POST['role']);
    
    // Validasi
    if (empty($nrp) || empty($password) || empty($confirm_password)) {
        $error = "Semua field harus diisi!";
    } elseif (!ctype_digit($nrp)) {
        $error = "NRP harus berupa angka!";
    } elseif (strlen($nrp) < 6) { 
        $error = "NRP minimal 6 digit!";
    } elseif ($password !== $confirm_password) {
        $error = "Password dan konfirmasi password tidak sama!";
    } elseif (strlen($password) < 6) {
        $error = "Password minimal 6 karakter!";
    } else {
        // Cek apakah NRP sudah ada
        $check_query = "SELECT * FROM users WHERE nrp = ?";
        $stmt = mysqli_prepare($conn, $check_query);
        mysqli_stmt_bind_param($stmt, "s", $nrp);
        mysqli_stmt_execute($stmt);
        $result = mysqli_stmt_get_result($stmt);
        
        if (mysqli_num_rows($result) > 0) {
            $error = "NRP sudah terdaftar!";
        } else {
            // Hash password
            $hashed_password = password_hash($password, PASSWORD_DEFAULT);
            
            // Insert ke database
            $insert_query = "INSERT INTO users (nrp, password, role) VALUES (?, ?, ?)";
            $stmt2 = mysqli_prepare($conn, $insert_query);
            mysqli_stmt_bind_param($stmt2, "sss", $nrp, $hashed_password, $role);
            
            if (mysqli_stmt_execute($stmt2)) {
                $user_id = mysqli_insert_id($conn);
                catat_log($_SESSION['user_id'], 'TAMBAH USER', 'Admin menambahkan user baru dengan NRP: ' . $nrp);
                $success = "User berhasil ditambahkan! NRP: $nrp";
                
                // Reset form (kecuali token)
                $nrp = '';
                $password = '';
            } else {
                $error = "Gagal menambahkan user: " . mysqli_error($conn);
            }
            mysqli_stmt_close($stmt2);
        }
        mysqli_stmt_close($stmt);
    }
}
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tambah User - KORPRAPORT</title>
    <link rel="stylesheet" href="../assets/css/masterpersonel.css">
    <link rel="stylesheet" href="../assets/css/adduser.css">
    <style>
        .main-content {
            padding-top: 30px;
        }
        .card {
            max-width: 600px;
            margin: 0 auto;
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
            <li><a href="dashboard.php"><span>📊</span> Dashboard</a></li>
            <li><a href="masterpersonel.php"><span>👥</span> Master Data Personel</a></li>
            <li><a href="adduser.php" class="active"><span>➕</span> Tambah User</a></li>
            <li><a href="historylog.php"><span>📋</span> History Log</a></li>
            <li><a href="../auth/logout.php"><span>🚪</span> Logout</a></li>
        </ul>
    </div>
    
    <div class="main-content">
        <div class="top-bar">
            <h1>Tambah User Baru</h1>
            <div class="user-info">
                 <a href="masterpersonel.php" class="btn btn-secondary">Kembali ke Master Data</a>
            </div>
        </div>

        <div class="card">
            <div class="form-info">
                ℹ️ User yang ditambahkan akan login menggunakan NRP dan password yang Anda buat. 
                Setelah login pertama kali, user akan diminta melengkapi data diri.
            </div>
            
            <?php if ($success): ?>
                <div class="alert alert-success"><?php echo $success; ?></div>
            <?php endif; ?>
            
            <?php if ($error): ?>
                <div class="alert alert-error"><?php echo $error; ?></div>
            <?php endif; ?>
            
            <form method="POST" action="">
                <!-- SECURITY FIX: Tambahkan CSRF Token -->
                <input type="hidden" name="csrf_token" value="<?php echo generate_csrf_token(); ?>">

                <div class="form-group">
                    <label for="nrp">NRP *</label>
                    <input type="text" id="nrp" name="nrp" required 
                           placeholder="Masukkan NRP (minimal 6 digit angka)" 
                           pattern="[0-9]{6,}"
                           value="<?php echo isset($nrp) ? htmlspecialchars($nrp) : ''; ?>">
                    <small style="color: #666;">Contoh: 000001</small>
                </div>
                
                <div class="form-group">
                    <label for="password">Password *</label>
                    <input type="password" id="password" name="password" required 
                           placeholder="Minimal 6 karakter">
                </div>
                
                <div class="form-group">
                    <label for="confirm_password">Konfirmasi Password *</label>
                    <input type="password" id="confirm_password" name="confirm_password" required 
                           placeholder="Ketik ulang password">
                </div>
                
                <div class="form-group">
                    <label for="role">Role *</label>
                    <select id="role" name="role" required>
                        <option value="user" <?php echo (isset($_POST['role']) && $_POST['role'] == 'user') ? 'selected' : ''; ?>>User</option>
                        <option value="admin" <?php echo (isset($_POST['role']) && $_POST['role'] == 'admin') ? 'selected' : ''; ?>>Admin</option>
                    </select>
                </div>
                
                <button type="submit" class="btn btn-primary">Tambah User</button>
            </form>
        </div>
    </div>
    
    <script src="../assets/js/adduser.js"></script>
</body>
</html>