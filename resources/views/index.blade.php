<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Daily Data Query</title>
    <link rel="stylesheet" href="./index.css">
    <!-- Sử dụng Bootstrap CSS từ CDN -->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <!-- Sử dụng thư viện SweetAlert2 cho cửa sổ thông báo đẹp -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/sweetalert2@10">
    <!-- CSS tùy chỉnh cho trang web -->
    <style>
        body {
            padding: 20px;
        }
    </style>
</head>
<body>

    <div class="container">
        <h1 class="mt-5">Data Query</h1>

        <!-- Button để kích hoạt yêu cầu -->
        <button class="btn btn-primary mt-3" id="checkDatabaseBtn">Check Database</button>

        <!-- Container để hiển thị kết quả từ yêu cầu Ajax -->
        <div id="resultContainer" class="mt-3"></div>
    </div>

    <!-- Sử dụng Bootstrap JS và Popper.js từ CDN -->
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.9.2/dist/umd/popper.min.js"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
    <!-- Sử dụng SweetAlert2 JS từ CDN -->
    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@10"></script>
    <!-- JavaScript tùy chỉnh cho trang web -->
    <script>
        // Sự kiện click cho nút kiểm tra cơ sở dữ liệu
        document.getElementById('checkDatabaseBtn').addEventListener('click', function() {
    fetch('/check-database', {
        method: 'POST', // Sử dụng phương thức POST
    })
    .then(response => response.text())
    .then(data => {
        document.getElementById('resultContainer').innerHTML = data;
    })
    .catch(error => {
        console.error('Fetch Error:', error);
        document.getElementById('resultContainer').innerHTML = '<div class="alert alert-danger" role="alert">Error checking database.</div>';
    });
});


    </script>
</body>
</html>
