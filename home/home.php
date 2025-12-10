<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Korp-Report</title>
    <link rel="stylesheet" href="../assets/css/home.css">
</head>
<body>
    <div class="dashboard">
        <img src="../assets/img/logo.png" alt="logo" class="logo">
        <h1 class="title">PUSINFOLAHTA<BR>KORP REPORT</h1>
            <button class="login-btn" onclick="handleLogin()">MASUK</button>
    </div>

    <script>
function handleLogin() {
    window.location.href = "../auth/login.php";
}
</script>
</body>
</html>