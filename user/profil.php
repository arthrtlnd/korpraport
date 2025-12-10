<?php
session_start();
require_once '../app/koneksi.php';
check_user();

$nrp = $_SESSION['nrp'];
$success = '';
$error = '';
$first_login = isset($_GET['first_login']) && $_GET['first_login'] == '1';

if (isset($_SESSION['success_message'])) {
    $success = $_SESSION['success_message'];
    unset($_SESSION['success_message']); 
}

// --- LOGIKA DATABASE ---
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
          WHERE p.nrp = ?";

$stmt = mysqli_prepare($conn, $query);
mysqli_stmt_bind_param($stmt, "s", $nrp);
mysqli_stmt_execute($stmt);
$result = mysqli_stmt_get_result($stmt);
$personel = mysqli_fetch_assoc($result);

$is_new = ($personel == null);
$edit_mode = $is_new; 

// Dropdown data
$korp_list = mysqli_query($conn, "SELECT * FROM korp ORDER BY SEBUTAN");
// REVISI: Urutan pangkat DESC (Jenderal paling atas)
$pangkat_list = mysqli_query($conn, "SELECT * FROM pangkat ORDER BY kd_pkt DESC");
$matra_list = mysqli_query($conn, "SELECT * FROM matra ORDER BY MTR");
$satker_list = mysqli_query($conn, "SELECT * FROM satker ORDER BY nama_satker");
$gender_list = mysqli_query($conn, "SELECT * FROM gender");

// --- HANDLE UPLOAD FOTO ---
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['foto_profil'])) {
    verify_csrf_token($_POST['csrf_token'] ?? '');
    $file = $_FILES['foto_profil'];
    if ($file['error'] == 0) {
        $allowed_ext = ['jpg', 'jpeg', 'png', 'gif'];
        $ext = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
        if (in_array($ext, $allowed_ext) && $file['size'] <= 2000000) {
            $finfo = finfo_open(FILEINFO_MIME_TYPE);
            $mime = finfo_file($finfo, $file['tmp_name']);
            finfo_close($finfo);
            $allowed_mime = ['image/jpeg', 'image/png', 'image/gif'];
            if (in_array($mime, $allowed_mime)) {
                $upload_dir = '../uploads/profile/';
                if (!file_exists($upload_dir)) mkdir($upload_dir, 0777, true);
                $new_filename = $nrp . '_' . bin2hex(random_bytes(8)) . '.' . $ext;
                $upload_path = $upload_dir . $new_filename;
                if (move_uploaded_file($file['tmp_name'], $upload_path)) {
                    if (!$is_new && !empty($personel['foto_profil']) && file_exists($upload_dir . $personel['foto_profil'])) unlink($upload_dir . $personel['foto_profil']);
                    $stmt2 = mysqli_prepare($conn, "UPDATE personel SET foto_profil = ? WHERE nrp = ?");
                    mysqli_stmt_bind_param($stmt2, "ss", $new_filename, $nrp);
                    mysqli_stmt_execute($stmt2);
                    $_SESSION['success_message'] = "Foto profil berhasil diupdate!";
                    header("Location: " . $_SERVER['PHP_SELF'] . ($first_login ? "?first_login=1" : ""));
                    exit;
                }
            } else { $error = "Format file tidak valid!"; }
        } else { $error = "File maksimal 2MB dan harus format gambar!"; }
    }
}

// --- HANDLE SAVE DATA ---
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['save_data'])) {
    verify_csrf_token($_POST['csrf_token'] ?? '');
    // Ambil input
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
    $no_sprint_lama = !empty($_POST['no_sprint_lama']) ? clean_input($_POST['no_sprint_lama']) : null;

    if (empty($nama) || empty($tempat_lahir) || empty($tanggal_lahir)) {
        $error = "Nama, Tempat Lahir, dan Tanggal Lahir wajib diisi!";
        $edit_mode = true;
    } else {
        if ($is_new) {
            $stmt_save = mysqli_prepare($conn, "INSERT INTO personel (nrp, nama, tempat_lahir, tanggal_lahir, korp, pangkat, matra, kd_satker, alamat, nik, no_hp, no_kep, no_sprint, kd_gender, satker_lama, no_kep_lama, no_sprint_lama) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            mysqli_stmt_bind_param($stmt_save, "sssssssssssssssss", $nrp, $nama, $tempat_lahir, $tanggal_lahir, $korp, $pangkat, $matra, $kd_satker, $alamat, $nik, $no_hp, $no_kep, $no_sprint, $kd_gender, $satker_lama, $no_kep_lama, $no_sprint_lama);
        } else {
            $stmt_save = mysqli_prepare($conn, "UPDATE personel SET nama=?, tempat_lahir=?, tanggal_lahir=?, korp=?, pangkat=?, matra=?, kd_satker=?, alamat=?, nik=?, no_hp=?, no_kep=?, no_sprint=?, kd_gender=?, satker_lama=?, no_kep_lama=?, no_sprint_lama=? WHERE nrp=?");
            mysqli_stmt_bind_param($stmt_save, "sssssssssssssssss", $nama, $tempat_lahir, $tanggal_lahir, $korp, $pangkat, $matra, $kd_satker, $alamat, $nik, $no_hp, $no_kep, $no_sprint, $kd_gender, $satker_lama, $no_kep_lama, $no_sprint_lama, $nrp);
        }
        
        if (mysqli_stmt_execute($stmt_save)) {
            catat_log($_SESSION['user_id'], $is_new ? 'TAMBAH DATA DIRI' : 'UPDATE DATA DIRI', 'User ' . $nrp . ' update data');
            $_SESSION['success_message'] = "Data berhasil disimpan!";
            header("Location: " . $_SERVER['PHP_SELF']);
            exit;
        } else { $error = "Gagal menyimpan data!"; $edit_mode = true; }
    }
}

$foto_path = '../uploads/profile/' . ($personel['foto_profil'] ?? '');
if (empty($personel['foto_profil'] ?? true) || !file_exists($foto_path)) {
    $foto_path = 'https://via.placeholder.com/200?text=' . urlencode($is_new ? 'User' : substr($personel['nama'], 0, 1));
}
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Profil - KORPRAPORT</title>
    <link rel="stylesheet" href="../assets/css/profil.css">
    <style>
        .container { max-width: 900px; }
        .profile-header {
            background: transparent;
            box-shadow: none;
            padding: 20px 0;
            margin-bottom: 20px;
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 25px;
            text-align: left;
        }
        .profile-photo img {
            width: 120px;
            height: 120px;
            border-radius: 50%;
            object-fit: cover;
            border: 4px solid #7d1c1c;
            box-shadow: 0 4px 10px rgba(0,0,0,0.1);
        }
        .profile-info h2 {
            font-size: 22px;
            color: #333;
            margin-bottom: 5px;
            font-weight: 700;
            text-transform: uppercase;
        }
        .profile-info .subtitle {
            font-size: 15px;
            color: #666;
        }
        .card {
            background: white;
            border-radius: 10px;
            padding: 25px;
            box-shadow: 0 2px 15px rgba(0,0,0,0.05);
            margin-bottom: 25px;
        }
        .card-header-flex {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            border-bottom: none;
        }
        .section-title {
            font-size: 18px;
            color: #7d1c1c;
            font-weight: 700;
            margin: 0;
            border: none;
        }
        .btn-edit-custom {
            background-color: #FFA502;
            color: white;
            border: none;
            padding: 8px 20px;
            border-radius: 5px;
            font-weight: 600;
            font-size: 14px;
            cursor: pointer;
            text-decoration: none;
            transition: background 0.3s;
        }
        .btn-edit-custom:hover { background-color: #e69500; }
        .form-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-bottom: 15px;
        }
        .form-group { margin-bottom: 15px; }
        .form-group label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: #444;
            font-size: 13px;
        }
        .form-group input, .form-group select, .form-group textarea {
            width: 100%;
            padding: 10px 12px;
            border: 1px solid #ddd;
            border-radius: 5px;
            font-size: 14px;
            background-color: #fff;
            color: #333;
        }
        .form-readonly input, .form-readonly select, .form-readonly textarea {
            background-color: #f8f9fa;
            border-color: #e9ecef;
            pointer-events: none;
            color: #666;
        }
        .data-box {
            background-color: #fcfcfc;
            border: 1px solid #eee;
            border-radius: 8px;
            padding: 15px;
        }
        .data-box-title {
            font-size: 14px;
            color: #666;
            font-weight: 600;
            margin-bottom: 15px;
        }
        .upload-btn {
            position: absolute;
            bottom: 5px;
            right: 5px;
            background: #7d1c1c;
            width: 35px; height: 35px;
            border-radius: 50%;
            display: flex; align-items: center; justify-content: center;
            color: white; cursor: pointer;
            border: 2px solid white;
        }
        .action-buttons {
            margin-top: 20px;
            display: none;
        }
        .btn-primary {
            background: linear-gradient(135deg, #e33333 0%, #7d1c1c 100%);
            color: white; padding: 10px 25px; border: none; border-radius: 5px; cursor: pointer; font-weight: bold;
        }
        .btn-secondary {
            background: #ccc; color: #333; padding: 10px 25px; border: none; border-radius: 5px; cursor: pointer; font-weight: bold; margin-left: 10px;
        }
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
                <span>Selamat datang, <strong><?php echo $is_new ? 'User Baru' : explode(' ', $personel['nama'])[0]; ?></strong></span>
                <a href="../auth/logout.php">Keluar</a>
            </div>
        </div>
    </header>
    
    <div class="container">
        <?php if ($success): ?><div class="alert alert-success"><?php echo $success; ?></div><?php endif; ?>
        <?php if ($error): ?><div class="alert alert-error"><?php echo $error; ?></div><?php endif; ?>
        
        <div class="profile-header">
            <div class="profile-photo" style="position: relative;">
                <img src="<?php echo $foto_path; ?>" alt="Foto Profil" id="previewFoto">
                <form method="POST" enctype="multipart/form-data" id="uploadForm">
                    <input type="hidden" name="csrf_token" value="<?php echo generate_csrf_token(); ?>">
                    <label class="upload-btn" for="foto_profil" title="Ganti Foto">
                        📷
                        <input type="file" name="foto_profil" id="foto_profil" accept="image/*" onchange="document.getElementById('uploadForm').submit();" style="display:none;">
                    </label>
                </form>
            </div>
            <div class="profile-info">
                <h2><?php echo $is_new ? 'LENGKAPI DATA' : $personel['nama']; ?></h2>
                <div class="subtitle">
                    <?php if(!$is_new): ?>
                        <?php echo $personel['pangkat_sebutan']; ?> (<?php echo $personel['korp_sebutan']; ?>) - NRP: <?php echo $nrp; ?>
                    <?php else: ?>
                        NRP: <?php echo $nrp; ?>
                    <?php endif; ?>
                </div>
            </div>
        </div>
        
        <form method="POST" action="" id="profileForm" class="<?php echo $edit_mode ? '' : 'form-readonly'; ?>">
            <input type="hidden" name="csrf_token" value="<?php echo generate_csrf_token(); ?>">

            <!-- DATA KEPANGKATAN CARD -->
            <div class="card">
                <div class="card-header-flex">
                    <div class="section-title">Data Kepangkatan</div>
                    <?php if (!$edit_mode): ?>
                        <button type="button" class="btn-edit-custom" onclick="enableEdit()">Edit</button>
                    <?php endif; ?>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>Pangkat</label>
                        <select name="pangkat" <?php echo !$edit_mode ? 'disabled' : ''; ?>>
                            <option value="">-- Pilih Pangkat --</option>
                            <?php mysqli_data_seek($pangkat_list, 0); while($row = mysqli_fetch_assoc($pangkat_list)): ?>
                                <option value="<?php echo $row['kd_pkt']; ?>" <?php echo (isset($personel['pangkat']) && $personel['pangkat'] == $row['kd_pkt']) ? 'selected' : ''; ?>><?php echo $row['sebutan']; ?></option>
                            <?php endwhile; ?>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Korp</label>
                        <select name="korp" id="korp" <?php echo !$edit_mode ? 'disabled' : ''; ?>>
                            <option value="">-- Pilih Korp --</option>
                            <?php mysqli_data_seek($korp_list, 0); while($row = mysqli_fetch_assoc($korp_list)): ?>
                                <option value="<?php echo $row['KORPSID']; ?>" <?php echo (isset($personel['korp']) && $personel['korp'] == $row['KORPSID']) ? 'selected' : ''; ?>><?php echo $row['SEBUTAN']; ?></option>
                            <?php endwhile; ?>
                        </select>
                    </div>
                </div>
                
                <div class="form-group">
                    <label>Matra</label>
                    <select name="matra" id="matra" <?php echo !$edit_mode ? 'disabled' : ''; ?>>
                        <option value="">-- Pilih Matra --</option>
                        <?php mysqli_data_seek($matra_list, 0); while($row = mysqli_fetch_assoc($matra_list)): ?>
                            <option value="<?php echo $row['MTR']; ?>" <?php echo (isset($personel['matra']) && $personel['matra'] == $row['MTR']) ? 'selected' : ''; ?>><?php echo $row['Nama']; ?></option>
                        <?php endwhile; ?>
                    </select>
                </div>

                <!-- DATA LAMA & BARU -->
                <div class="form-row">
                    <!-- Box Data Lama -->
                    <div class="data-box">
                        <div class="data-box-title">Data Lama</div>
                        <div class="form-group">
                            <label>Satuan Kerja Lama</label>
                            <select name="satker_lama" <?php echo !$edit_mode ? 'disabled' : ''; ?>>
                                <option value="">-- Pilih Satker --</option>
                                <?php mysqli_data_seek($satker_list, 0); while($row = mysqli_fetch_assoc($satker_list)): ?>
                                    <option value="<?php echo $row['kd_satker']; ?>" <?php echo (isset($personel['satker_lama']) && $personel['satker_lama'] == $row['kd_satker']) ? 'selected' : ''; ?>><?php echo $row['nama_satker']; ?></option>
                                <?php endwhile; ?>
                            </select>
                        </div>
                        <div class="form-group">
                            <label>No. KEP Lama</label>
                            <input type="text" name="no_kep_lama" value="<?php echo $personel['no_kep_lama'] ?? ''; ?>" <?php echo !$edit_mode ? 'readonly' : ''; ?> placeholder="Opsional">
                        </div>
                        <div class="form-group">
                            <label>No. Sprint Lama</label>
                            <input type="text" name="no_sprint_lama" value="<?php echo $personel['no_sprint_lama'] ?? ''; ?>" <?php echo !$edit_mode ? 'readonly' : ''; ?> placeholder="Opsional">
                        </div>
                    </div>
                    
                    <!-- Box Data Baru -->
                    <div class="data-box">
                        <div class="data-box-title">Data Baru</div>
                        <div class="form-group">
                            <label>Satuan Kerja Baru</label>
                            <select name="kd_satker" <?php echo !$edit_mode ? 'disabled' : ''; ?>>
                                <option value="">-- Pilih Satker --</option>
                                <?php mysqli_data_seek($satker_list, 0); while($row = mysqli_fetch_assoc($satker_list)): ?>
                                    <option value="<?php echo $row['kd_satker']; ?>" <?php echo (isset($personel['kd_satker']) && $personel['kd_satker'] == $row['kd_satker']) ? 'selected' : ''; ?>><?php echo $row['nama_satker']; ?></option>
                                <?php endwhile; ?>
                            </select>
                        </div>
                        <div class="form-group">
                            <label>No. KEP Baru</label>
                            <input type="text" name="no_kep" value="<?php echo $personel['no_kep'] ?? ''; ?>" <?php echo !$edit_mode ? 'readonly' : ''; ?> placeholder="Opsional">
                        </div>
                        <div class="form-group">
                            <label>No. Sprint Baru</label>
                            <input type="text" name="no_sprint" value="<?php echo $personel['no_sprint'] ?? ''; ?>" <?php echo !$edit_mode ? 'readonly' : ''; ?> placeholder="Opsional">
                        </div>
                    </div>
                </div>
            </div>

            <!-- DATA PRIBADI CARD -->
            <div class="card">
                <div class="card-header-flex">
                    <div class="section-title">Data Pribadi</div>
                </div>
                
                <div class="form-group">
                    <label>Nama Lengkap</label>
                    <input type="text" name="nama" value="<?php echo $personel['nama'] ?? ''; ?>" <?php echo $edit_mode ? 'required' : 'readonly'; ?>>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>NIK</label>
                        <input type="text" name="nik" value="<?php echo $personel['nik'] ?? ''; ?>" <?php echo $edit_mode ? '' : 'readonly'; ?>>
                    </div>
                    <div class="form-group">
                        <label>Jenis Kelamin</label>
                        <select name="kd_gender" <?php echo !$edit_mode ? 'disabled' : ''; ?>>
                            <option value="">-- Pilih --</option>
                            <?php mysqli_data_seek($gender_list, 0); while($row = mysqli_fetch_assoc($gender_list)): ?>
                                <option value="<?php echo $row['kd_gender']; ?>" <?php echo (isset($personel['kd_gender']) && $personel['kd_gender'] == $row['kd_gender']) ? 'selected' : ''; ?>><?php echo $row['gender']; ?></option>
                            <?php endwhile; ?>
                        </select>
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>Tanggal Lahir</label>
                        <input type="date" name="tanggal_lahir" value="<?php echo $personel['tanggal_lahir'] ?? ''; ?>" <?php echo $edit_mode ? 'required' : 'readonly'; ?>>
                    </div>
                    <div class="form-group">
                        <label>Tempat Lahir</label>
                        <input type="text" name="tempat_lahir" value="<?php echo $personel['tempat_lahir'] ?? ''; ?>" <?php echo $edit_mode ? 'required' : 'readonly'; ?>>
                    </div>
                </div>
                
                <div class="form-group">
                    <label>No. HP</label>
                    <input type="text" name="no_hp" value="<?php echo $personel['no_hp'] ?? ''; ?>" <?php echo !$edit_mode ? 'readonly' : ''; ?>>
                </div>
            </div>
            
            <!-- ALAMAT CARD -->
            <div class="card">
                <div class="card-header-flex">
                    <div class="section-title">Alamat</div>
                </div>
                <div class="form-group">
                    <label>Alamat Lengkap</label>
                    <textarea name="alamat" rows="3" <?php echo !$edit_mode ? 'readonly' : ''; ?>><?php echo $personel['alamat'] ?? ''; ?></textarea>
                </div>
            </div>
            
            <div class="action-buttons" id="actionButtons" style="display: <?php echo $edit_mode ? 'block' : 'none'; ?>; margin-bottom: 50px;">
                <button type="submit" name="save_data" class="btn-primary">Simpan</button>
                <?php if(!$is_new): ?>
                    <button type="button" class="btn-secondary" onclick="location.reload()">Batal</button>
                <?php endif; ?>
            </div>
        </form>
    </div>
    
    <script src="../assets/js/profil.js"></script>
    <script>
        function enableEdit() {
            document.getElementById('profileForm').classList.remove('form-readonly');
            document.querySelectorAll('#profileForm input, #profileForm select, #profileForm textarea').forEach(el => {
                if(el.name !== 'csrf_token') { el.removeAttribute('readonly'); el.removeAttribute('disabled'); }
            });
            document.getElementById('actionButtons').style.display = 'block';
            document.querySelector('.btn-edit-custom').style.display = 'none';
            
            if(typeof filterKorp === "function") filterKorp();
        }

        const korpByMatra = {'1': ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'K1', 'M1', 'N1', 'P1', 'Q1', 'R1', 'X1', 'Y1', 'Z1', 'A3'], '2': ['12', '22', '32', '42', '52', '62', '72', '82'], '3': ['13', '23', '33', '43', '53', '63', '73', '83', '93', 'A3'], '0': []};
        const matraSelect = document.getElementById('matra');
        const korpSelect = document.getElementById('korp');
        const allKorpOptions = Array.from(korpSelect.options);
        
        function filterKorp() {
            const selectedMatra = matraSelect.value;
            const currentKorp = korpSelect.value;
            const isReadOnly = document.getElementById('profileForm').classList.contains('form-readonly');
            
            // Ambil elemen select pangkat
            const pangkatSelect = document.querySelector('select[name="pangkat"]');

            korpSelect.innerHTML = '<option value="">-- Pilih Korp --</option>';
            
            if (selectedMatra === '0') { 
                // Jika PNS, disable Korp DAN Pangkat
                korpSelect.disabled = true; 
                korpSelect.value = ''; 
                
                if(pangkatSelect) {
                    pangkatSelect.disabled = true;
                    pangkatSelect.value = '';
                }
            } 
            else if (selectedMatra && korpByMatra[selectedMatra]) {
                korpSelect.disabled = isReadOnly;
                
                if(pangkatSelect) {
                    pangkatSelect.disabled = isReadOnly;
                }

                allKorpOptions.forEach(option => { if (option.value && korpByMatra[selectedMatra].includes(option.value)) { const newOption = option.cloneNode(true); if (option.value === currentKorp) newOption.selected = true; korpSelect.appendChild(newOption); }});
            } else {
                korpSelect.disabled = isReadOnly;
                
                if(pangkatSelect) {
                    pangkatSelect.disabled = isReadOnly;
                }

                allKorpOptions.forEach(option => { if (option.value) { const newOption = option.cloneNode(true); if (option.value === currentKorp) newOption.selected = true; korpSelect.appendChild(newOption); }});
            }
        }
        if (matraSelect) { matraSelect.addEventListener('change', filterKorp); window.addEventListener('DOMContentLoaded', filterKorp); }
    </script>
</body>
</html>