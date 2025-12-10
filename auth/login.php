<?php
session_start();
require_once '../app/koneksi.php';

// Jika sudah login, redirect sesuai role
if (isset($_SESSION['login']) && $_SESSION['login'] === true) {
    if ($_SESSION['role'] === 'admin') {
        header("Location: ../admin/dashboard.php");
    } else {
        header("Location: ../user/profil.php");
    }
    exit();
}

$error = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $nrp = clean_input($_POST['nrp']);
    $password = $_POST['password'];
    
    // Validasi input kosong
    if (empty($nrp) || empty($password)) {
        $error = "NRP dan Password harus diisi!";
    } else {
        // Query user
        $query = "SELECT * FROM users WHERE nrp = ?";
        $stmt = mysqli_prepare($conn, $query);
        mysqli_stmt_bind_param($stmt, "s", $nrp);
        mysqli_stmt_execute($stmt);
        $result = mysqli_stmt_get_result($stmt);
        
        if ($row = mysqli_fetch_assoc($result)) {
            // Verifikasi password
            if (password_verify($password, $row['password'])) {
                // Set session
                $_SESSION['login'] = true;
                $_SESSION['user_id'] = $row['id'];
                $_SESSION['nrp'] = $row['nrp'];
                $_SESSION['role'] = $row['role'];
                
                // Catat log login
                catat_log($row['id'], 'LOGIN', 'User ' . $nrp . ' berhasil login');
                
                // Redirect sesuai role
                if ($row['role'] === 'admin') {
                    header("Location: ../admin/dashboard.php");
                } else {
                    // Cek apakah user sudah melengkapi data
                    $check_personel = "SELECT id FROM personel WHERE nrp = ?";
                    $stmt2 = mysqli_prepare($conn, $check_personel);
                    mysqli_stmt_bind_param($stmt2, "s", $nrp);
                    mysqli_stmt_execute($stmt2);
                    $result2 = mysqli_stmt_get_result($stmt2);
                    
                    // REVISI 5: Ubah alur login user baru
                    if (mysqli_num_rows($result2) == 0) {
                        // Belum ada data, redirect ke profil.php dengan flag
                        header("Location: ../user/profil.php?first_login=1");
                    } else {
                        // Sudah ada data, redirect ke profil.php
                        header("Location: ../user/profil.php");
                    }
                }
                exit();
            } else {
                $error = "Password salah!";
            }
        } else {
            $error = "NRP tidak ditemukan!";
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
    <title>Login - KORPRAPORT</title>
    <link rel="stylesheet" href="../assets/css/login.css">
</head>
<body>
    <header>
        <div class="header">
        <img src="../assets/img/logo-2.png" alt="logo">
        <div class="header-title">
            <h2>SISTEM INFORMASI PERSONEL</h2>
            <h2>MARKAS BESAR TENTARA NASIONAL INDONESIA</h2>
        </div>
    </header>
    <div class="container">
        <div class="left-section">
            <div class="logo-section">
                <div class="main-logo"><img src="../assets/img/logo.png" alt="main-logo"></div>
                <div class="divider"></div>
                <div class="brand-text">PUSINFO<br>LAHTA</div>
            </div>
        </div>
        <div class="right-section">
            <div class="decorative-wave"></div>
            <div class="login-card">
                <div class="login-title">Masuk ke dalam</div>
                <div class="login-subtitle">Akun Anda</div>

                <?php if ($error): ?>
            <div class="error-message"><?php echo $error; ?></div>
        <?php endif; ?>
                
                <form method="post">
                    <div class="form-group">
                        <label class="form-label">NRP</label>
                        <div class="input-wrapper">
                            <input type="text" name="nrp" class="form-input" required>
                        </div>
                    </div>
                    <div class="form-group">
                        <label class="form-label">Kata sandi</label>
                        <div class="input-wrapper">
                            <input type="password" name="password" class="form-input" required>
                        </div>
                    </div>
                    <button type="submit" name="login" class="login-button">Masuk</button>
                    <ul>
                        <li><a href="forgot-pw.php">Lupa Password</a></li>
                    </ul>
                </form>

            </div>
        </div>
    </div>
</body>
</html>