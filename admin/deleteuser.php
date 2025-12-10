<?php
session_start();
require_once '../app/koneksi.php';
check_admin();

if (!isset($_GET['id'])) {
    header("Location: masterpersonel.php");
    exit();
}

$id = clean_input($_GET['id']);

// SECURITY FIX: Gunakan Transaction Database (ACID)
mysqli_begin_transaction($conn);

try {
    // 1. Ambil NRP dan Foto dari tabel personel
    $query = "SELECT nrp, foto_profil FROM personel WHERE id = ?";
    $stmt = mysqli_prepare($conn, $query);
    mysqli_stmt_bind_param($stmt, "i", $id);
    mysqli_stmt_execute($stmt);
    $result = mysqli_stmt_get_result($stmt);
    $data = mysqli_fetch_assoc($result);

    if (!$data) {
        throw new Exception("Data personel tidak ditemukan.");
    }

    $nrp = $data['nrp'];
    $foto_profil = $data['foto_profil'];

    // 2. Ambil ID User dari tabel users (PENTING untuk hapus history)
    $query_user = "SELECT id FROM users WHERE nrp = ?";
    $stmt_u = mysqli_prepare($conn, $query_user);
    mysqli_stmt_bind_param($stmt_u, "s", $nrp);
    mysqli_stmt_execute($stmt_u);
    $result_u = mysqli_stmt_get_result($stmt_u);
    $user_data = mysqli_fetch_assoc($result_u);

    // 3. Hapus Data History Log (SOLUSI BUG: Hapus history dulu agar tidak nyangkut)
    if ($user_data) {
        $user_id = $user_data['id'];
        $stmt_log = mysqli_prepare($conn, "DELETE FROM history_log WHERE user_id = ?");
        mysqli_stmt_bind_param($stmt_log, "i", $user_id);
        mysqli_stmt_execute($stmt_log);
        mysqli_stmt_close($stmt_log);
    }

    // 4. Hapus data personel
    $stmt1 = mysqli_prepare($conn, "DELETE FROM personel WHERE id = ?");
    mysqli_stmt_bind_param($stmt1, "i", $id);
    if (!mysqli_stmt_execute($stmt1)) {
        throw new Exception("Gagal menghapus data personel.");
    }
    mysqli_stmt_close($stmt1);

    // 5. Hapus data user login
    $stmt2 = mysqli_prepare($conn, "DELETE FROM users WHERE nrp = ?");
    mysqli_stmt_bind_param($stmt2, "s", $nrp);
    if (!mysqli_stmt_execute($stmt2)) {
        throw new Exception("Gagal menghapus user login.");
    }
    mysqli_stmt_close($stmt2);

    // Jika semua sukses, simpan perubahan permanen
    mysqli_commit($conn);

    // 6. Hapus file fisik foto
    if (!empty($foto_profil)) {
        $foto_path = '../uploads/profile/' . $foto_profil;
        if (file_exists($foto_path)) {
            unlink($foto_path);
        }
    }

    // Set pesan sukses untuk ditampilkan di halaman master
    $_SESSION['success'] = "Data personel NRP $nrp berhasil dihapus.";

} catch (Exception $e) {
    // Jika error, batalkan semua perubahan
    mysqli_rollback($conn);
    $_SESSION['error'] = "Gagal menghapus: " . $e->getMessage();
}

// Redirect kembali ke tabel master personel
header("Location: masterpersonel.php");
exit();
?>