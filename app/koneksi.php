<?php
// Konfigurasi Database
define('DB_HOST', 'localhost');
define('DB_USER', 'root');
define('DB_PASS', '');
define('DB_NAME', 'korpraport');

// Membuat koneksi
$conn = mysqli_connect(DB_HOST, DB_USER, DB_PASS, DB_NAME);

// Cek koneksi
if (!$conn) {
    die("Koneksi gagal: " . mysqli_connect_error());
}

// Set charset
mysqli_set_charset($conn, "utf8mb4");

// Fungsi untuk mencegah SQL Injection
function escape_string($data) {
    global $conn;
    return mysqli_real_escape_string($conn, $data);
}

// Fungsi untuk sanitasi input
function clean_input($data) {
    if ($data === null || $data === '') {
        return null;
    }
    $data = trim($data);
    $data = stripslashes($data);
    $data = htmlspecialchars($data);
    return $data;
}

// Fungsi catat history log
function catat_log($user_id, $aksi, $keterangan) {
    global $conn;
    $query = "INSERT INTO history_log (user_id, aksi, keterangan) VALUES (?, ?, ?)";
    $stmt = mysqli_prepare($conn, $query);
    mysqli_stmt_bind_param($stmt, "iss", $user_id, $aksi, $keterangan);
    mysqli_stmt_execute($stmt);
    mysqli_stmt_close($stmt);
}

// Fungsi cek login
function check_login() {
    if (!isset($_SESSION['login']) || $_SESSION['login'] !== true) {
        header("Location: ../auth/login.php");
        exit();
    }
}

// Fungsi cek admin
function check_admin() {
    check_login();
    if ($_SESSION['role'] !== 'admin') {
        header("Location: ../user/profil.php");
        exit();
    }
}

// Fungsi cek user
function check_user() {
    check_login();
    if ($_SESSION['role'] !== 'user') {
        header("Location: ../admin/dashboard.php");
        exit();
    }
}

// --- SECURITY HELPER (CSRF) ---
function generate_csrf_token() {
    if (empty($_SESSION['csrf_token'])) {
        $_SESSION['csrf_token'] = bin2hex(random_bytes(32));
    }
    return $_SESSION['csrf_token'];
}

function verify_csrf_token($token) {
    if (!isset($_SESSION['csrf_token']) || $token !== $_SESSION['csrf_token']) {
        die("Akses ditolak: Token keamanan tidak valid (CSRF Detected). Silakan refresh halaman.");
    }
    return true;
}
?>