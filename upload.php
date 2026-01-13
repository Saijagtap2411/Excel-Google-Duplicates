<!DOCTYPE html>
<html>
<head>
    <title>Excel Duplicate Checker</title>
</head>
<body>

<h2>Upload Excel File</h2>

<form action="process.php" method="POST" enctype="multipart/form-data">
    <input type="file" name="excel" accept=".xls,.xlsx" required>
    <br><br>
    <button type="submit">Check Duplicates</button>
</form>

</body>
</html>
