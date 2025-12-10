<?php
session_start();
require_once '../app/koneksi.php';

// Catat log logout jika user sudah login
if (isset($_SESSION['user_id']) && isset($_SESSION['nrp'])) {
    catat_log($_SESSION['user_id'], 'LOGOUT', 'User ' . $_SESSION['nrp'] . ' logout');
}

// Hapus semua session
session_unset();
session_destroy();

// Redirect ke halaman login
header("Location: login.php");
exit();
?>