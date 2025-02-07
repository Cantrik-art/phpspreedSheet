<!DOCTYPE html>
<html lang="id">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Upload & Download Excel</title>
</head>

<body>

    <h2>Upload File Excel</h2>

    <?php if (isset($_SESSION['message'])): ?>
        <p><?php echo $_SESSION['message']; ?></p>
        <?php unset($_SESSION['message']); // Hapus setelah ditampilkan 
        ?>
    <?php endif; ?>


    <form action="<?php echo base_url('excel/import'); ?>" method="post" enctype="multipart/form-data">
        <input type="file" name="file" required>
        <button type="submit">Upload</button>
    </form>

    <h2>Download Data sebagai Excel</h2>
    <a href="<?php echo base_url('excel/export'); ?>"><button>Download Excel</button></a>

</body>

</html>