<?php
session_start();
require_once '../app/koneksi.php';
check_user();

$nrp = $_SESSION['nrp'];
$first_login = isset($_GET['first_login']) ? true : false;

$success = '';
$error = '';

$check_query = "SELECT * FROM personel WHERE nrp = ?";
$stmt_check = mysqli_prepare($conn, $check_query);
mysqli_stmt_bind_param($stmt_check, "s", $nrp);
mysqli_stmt_execute($stmt_check);
$result_check = mysqli_stmt_get_result($stmt_check);
$personel = mysqli_fetch_assoc($result_check);
$is_new = ($personel == null);

$korp_list = mysqli_query($conn, "SELECT * FROM korp ORDER BY SEBUTAN");
// REVISI: Urutan DESC
$pangkat_list = mysqli_query($conn, "SELECT * FROM pangkat ORDER BY kd_pkt DESC");
$matra_list = mysqli_query($conn, "SELECT * FROM matra ORDER BY MTR");
$satker_list = mysqli_query($conn, "SELECT * FROM satker ORDER BY nama_satker");
$gender_list = mysqli_query($conn, "SELECT * FROM gender");

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $nama = clean_input($_POST['nama']);
    $tempat_lahir = clean_input($_POST['tempat_lahir']);
    $tanggal_lahir = clean_input($_POST['tanggal_lahir']);
    $korp = !empty($_POST['korp']) ? clean_input($_POST['korp']) : null;
    $pangkat = !empty($_POST['pangkat']) ? clean_input($_POST['pangkat']) : null;
    $matra = !empty($_POST['matra']) ? clean_input($_POST['matra']) : null;
    $kd_satker = !empty($_POST['kd_satker']) ? clean_input($_POST['kd_satker']) : null; 
    $alamat = !empty($_POST['alamat']) ? clean_input($_POST['alamat']) : null;
    $nik = !empty($_POST['nik']) ? clean_input($_POST['nik']) : null;
    $no_hp = !empty($_POST['no_hp']) ? clean_input($_POST['no_hp']) : null;
    $no_kep = !empty($_POST['no_kep']) ? clean_input($_POST['no_kep']) : null; 
    $no_sprint = !empty($_POST['no_sprint']) ? clean_input($_POST['no_sprint']) : null; 
    $kd_gender = !empty($_POST['kd_gender']) ? clean_input($_POST['kd_gender']) : null;
    
    $satker_lama = !empty($_POST['satker_lama']) ? clean_input($_POST['satker_lama']) : null;
    $no_kep_lama = !empty($_POST['no_kep_lama']) ? clean_input($_POST['no_kep_lama']) : null;
    // REVISI: Tambah simpan sprint lama
    $no_sprint_lama = !empty($_POST['no_sprint_lama']) ? clean_input($_POST['no_sprint_lama']) : null;

    if (empty($nama) || empty($tempat_lahir) || empty($tanggal_lahir)) {
        $error = "Nama, Tempat Lahir, dan Tanggal Lahir wajib diisi!";
    } else {
        if ($is_new) {
            $insert_query = "INSERT INTO personel (nrp, nama, tempat_lahir, tanggal_lahir, korp, pangkat, matra, kd_satker, alamat, nik, no_hp, no_kep, no_sprint, kd_gender, satker_lama, no_kep_lama, no_sprint_lama) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
            $stmt = mysqli_prepare($conn, $insert_query);
            mysqli_stmt_bind_param($stmt, "sssssssssssssssss", 
                $nrp, $nama, $tempat_lahir, $tanggal_lahir, $korp, $pangkat, 
                $matra, $kd_satker, $alamat, $nik, $no_hp, $no_kep, $no_sprint, $kd_gender,
                $satker_lama, $no_kep_lama, $no_sprint_lama);
            
            if (mysqli_stmt_execute($stmt)) {
                catat_log($_SESSION['user_id'], 'TAMBAH DATA DIRI', 'User ' . $nrp . ' melengkapi data diri');
                header("Location: profil.php");
                exit();
            } else {
                $error = "Gagal menyimpan data: " . mysqli_error($conn);
            }
        } else {
            $update_query = "UPDATE personel SET 
                            nama = ?, tempat_lahir = ?, tanggal_lahir = ?, 
                            korp = ?, pangkat = ?, matra = ?, kd_satker = ?,
                            alamat = ?, nik = ?, no_hp = ?, no_kep = ?, no_sprint = ?, kd_gender = ?,
                            satker_lama = ?, no_kep_lama = ?, no_sprint_lama = ?
                            WHERE nrp = ?";
            
            $stmt = mysqli_prepare($conn, $update_query);
            mysqli_stmt_bind_param($stmt, "sssssssssssssssss", 
                $nama, $tempat_lahir, $tanggal_lahir, $korp, $pangkat, 
                $matra, $kd_satker, $alamat, $nik, $no_hp, $no_kep, $no_sprint, $kd_gender,
                $satker_lama, $no_kep_lama, $no_sprint_lama, $nrp);
            
            if (mysqli_stmt_execute($stmt)) {
                catat_log($_SESSION['user_id'], 'UPDATE DATA DIRI', 'User ' . $nrp . ' mengupdate data diri');
                $success = "Data berhasil diupdate!";
                mysqli_stmt_execute($stmt_check);
                $result_check = mysqli_stmt_get_result($stmt_check);
                $personel = mysqli_fetch_assoc($result_check);
            } else {
                $error = "Gagal mengupdate data: " . mysqli_error($conn);
            }
        }
        if (isset($stmt)) { mysqli_stmt_close($stmt); }
    }
}
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title><?php echo $is_new ? 'Lengkapi Data' : 'Update Data'; ?> - KORPRAPORT</title>
    <link rel="stylesheet" href="../assets/css/update.css">
    <style>
        .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        .data-group { border: 1px solid #eee; border-radius: 8px; padding: 15px; margin-top: 10px; }
        .data-group-title { font-weight: 600; color: #555; margin-bottom: 15px; }
    </style>
</head>
<body>
    <header>
        <div class="header">
            <img src="../assets/img/logo-2.png" alt="logo">
                <div class="header-title">
                    <h2>SISTEM INFORMASI PERSONEL</h2>
                    <h2>MARKAS BESAR TENTARA NASIONAL INDONESIA</h2>
                </div>
                <div class="navbar-menu">
            <?php if (!$is_new): ?>
            <a href="profil.php">Kembali</a>
            <?php endif; ?>
            <a href="../auth/logout.php">Keluar</a>
        </div>
        </div>
    </header>
    
    <div class="container">
        <?php if ($first_login && $is_new): ?>
        <div class="welcome-box">
            <h3>🎉 Selamat Datang!</h3>
            <p>Ini adalah login pertama Anda. Silakan lengkapi data diri Anda di bawah ini.</p>
        </div>
        <?php endif; ?>
        
        <div class="card">
            <h2><?php echo $is_new ? 'Lengkapi Data Diri' : 'Update Data Diri'; ?></h2>
            <?php if ($success): ?><div class="alert alert-success"><?php echo $success; ?></div><?php endif; ?>
            <?php if ($error): ?><div class="alert alert-error"><?php echo $error; ?></div><?php endif; ?>
            
            <form method="POST" action="">
                <div class="section-title">Data Kepangkatan</div>
                <div class="form-row">
                    <div class="form-group">
                        <label for="pangkat">Pangkat</label>
                        <select id="pangkat" name="pangkat">
                            <option value="">-- Pilih Pangkat --</option>
                            <?php 
                            mysqli_data_seek($pangkat_list, 0);
                            while($row = mysqli_fetch_assoc($pangkat_list)): ?>
                                <option value="<?php echo $row['kd_pkt']; ?>"
                                    <?php echo (isset($personel['pangkat']) && $personel['pangkat'] == $row['kd_pkt']) ? 'selected' : ''; ?>>
                                    <?php echo $row['sebutan']; ?>
                                </option>
                            <?php endwhile; ?>
                        </select>
                    </div>
                    <div class="form-group">
                        <label for="korp">Korp</label>
                        <select id="korp" name="korp">
                            <option value="">-- Pilih Korp --</option>
                            <?php 
                            mysqli_data_seek($korp_list, 0);
                            while($row = mysqli_fetch_assoc($korp_list)): ?>
                                <option value="<?php echo $row['KORPSID']; ?>"
                                    <?php echo (isset($personel['korp']) && $personel['korp'] == $row['KORPSID']) ? 'selected' : ''; ?>>
                                    <?php echo $row['SEBUTAN']; ?>
                                </option>
                            <?php endwhile; ?>
                        </select>
                    </div>
                </div>
                
                <div class="form-group">
                    <label for="matra">Matra</label>
                    <select id="matra" name="matra">
                        <option value="">-- Pilih Matra --</option>
                        <?php 
                        mysqli_data_seek($matra_list, 0);
                        while($row = mysqli_fetch_assoc($matra_list)): ?>
                            <option value="<?php echo $row['MTR']; ?>"
                                <?php echo (isset($personel['matra']) && $personel['matra'] == $row['MTR']) ? 'selected' : ''; ?>>
                                <?php echo $row['Nama']; ?>
                            </option>
                        <?php endwhile; ?>
                    </select>
                </div>

                <div class="info-grid">
                    <div class="data-group">
                        <div class="data-group-title">Data Lama</div>
                        <div class="form-group">
                            <label>Satuan Kerja Lama</label>
                            <select id="satker_lama" name="satker_lama">
                                <option value="">-- Pilih Satker --</option>
                                <?php 
                                mysqli_data_seek($satker_list, 0);
                                while($row = mysqli_fetch_assoc($satker_list)): ?>
                                    <option value="<?php echo $row['kd_satker']; ?>"
                                        <?php echo (isset($personel['satker_lama']) && $personel['satker_lama'] == $row['kd_satker']) ? 'selected' : ''; ?>>
                                        <?php echo $row['nama_satker']; ?>
                                    </option>
                                <?php endwhile; ?>
                            </select>
                        </div>
                        <div class="form-group">
                            <label>No. KEP Lama</label>
                            <input type="text" id="no_kep_lama" name="no_kep_lama" maxlength="20" placeholder="Opsional" value="<?php echo $personel['no_kep_lama'] ?? ''; ?>">
                        </div>
                        <!-- REVISI: Tambah Input No Sprint Lama -->
                        <div class="form-group">
                            <label>No. Sprint Lama</label>
                            <input type="text" id="no_sprint_lama" name="no_sprint_lama" maxlength="20" placeholder="Opsional" value="<?php echo $personel['no_sprint_lama'] ?? ''; ?>">
                        </div>
                    </div>
                    
                    <div class="data-group">
                        <div class="data-group-title">Data Baru</div>
                        <div class="form-group">
                            <label>Satuan Kerja Baru</label>
                            <select id="kd_satker" name="kd_satker">
                                <option value="">-- Pilih Satker --</option>
                                <?php 
                                mysqli_data_seek($satker_list, 0);
                                while($row = mysqli_fetch_assoc($satker_list)): ?>
                                    <option value="<?php echo $row['kd_satker']; ?>"
                                        <?php echo (isset($personel['kd_satker']) && $personel['kd_satker'] == $row['kd_satker']) ? 'selected' : ''; ?>>
                                        <?php echo $row['nama_satker']; ?>
                                    </option>
                                <?php endwhile; ?>
                            </select>
                        </div>
                        <div class="form-group">
                            <label>No. KEP Baru</label>
                            <input type="text" id="no_kep" name="no_kep" maxlength="20"
                                   value="<?php echo $personel['no_kep'] ?? ''; ?>" placeholder="Opsional">
                        </div>
                        <!-- REVISI: Tambah Input No Sprint Baru -->
                        <div class="form-group">
                            <label>No. Sprint Baru</label>
                            <input type="text" id="no_sprint" name="no_sprint" maxlength="20"
                                   value="<?php echo $personel['no_sprint'] ?? ''; ?>" placeholder="Opsional">
                        </div>
                    </div>
                </div>

                <div class="section-title">Data Pribadi</div>
                <div class="form-group">
                    <label for="nama">Nama Lengkap <span class="required">*</span></label>
                    <input type="text" id="nama" name="nama" required value="<?php echo $personel['nama'] ?? ''; ?>">
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label for="nik">NIK</label>
                        <input type="text" id="nik" name="nik" maxlength="16" value="<?php echo $personel['nik'] ?? ''; ?>">
                    </div>
                    <div class="form-group">
                        <label for="kd_gender">Jenis Kelamin</label>
                        <select id="kd_gender" name="kd_gender">
                            <option value="">-- Pilih --</option>
                            <?php 
                            mysqli_data_seek($gender_list, 0);
                            while($row = mysqli_fetch_assoc($gender_list)): ?>
                                <option value="<?php echo $row['kd_gender']; ?>"
                                    <?php echo (isset($personel['kd_gender']) && $personel['kd_gender'] == $row['kd_gender']) ? 'selected' : ''; ?>>
                                    <?php echo $row['gender']; ?>
                                </option>
                            <?php endwhile; ?>
                        </select>
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label for="tempat_lahir">Tempat Lahir <span class="required">*</span></label>
                        <input type="text" id="tempat_lahir" name="tempat_lahir" required value="<?php echo $personel['tempat_lahir'] ?? ''; ?>">
                    </div>
                    <div class="form-group">
                        <label for="tanggal_lahir">Tanggal Lahir <span class="required">*</span></label>
                        <input type="date" id="tanggal_lahir" name="tanggal_lahir" required value="<?php echo $personel['tanggal_lahir'] ?? ''; ?>">
                    </div>
                </div>
                
                <div class="form-group">
                    <label for="no_hp">No. HP</label>
                    <input type="text" id="no_hp" name="no_hp" maxlength="15" value="<?php echo $personel['no_hp'] ?? ''; ?>">
                </div>
                                
                <div class="section-title">Alamat</div>
                <div class="form-group">
                    <label for="alamat">Alamat Lengkap</label>
                    <textarea id="alamat" name="alamat"><?php echo $personel['alamat'] ?? ''; ?></textarea>
                </div>
                
                <button type="submit" class="btn btn-primary">
                    <?php echo $is_new ? 'Simpan Data' : 'Update Data'; ?>
                </button>
                <?php if (!$is_new): ?>
                <a href="profil.php" class="btn btn-secondary">Batal</a>
                <?php endif; ?>
            </form>
        </div>
    </div>
    
    <script src="../assets/js/update.js"></script>
</body>
</html>