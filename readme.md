

# **PhpSpreadsheet Integration in CodeIgniter 3**

This project demonstrates how to **import and export Excel files** using **PhpSpreadsheet** in CodeIgniter 3. It follows the **MVC (Model-View-Controller)** architecture and integrates with a MySQL database.

---

## **1. Installation**

### **1.1 Install PhpSpreadsheet via Composer**
Run the following command in your project's root directory:
```bash
composer require phpoffice/phpspreadsheet
```

### **1.2 Enable Composer Autoload in CodeIgniter**
Modify `application/config/config.php`:
```php
$config['composer_autoload'] = FCPATH . 'vendor/autoload.php';
```

### **1.3 Enable Required Libraries and Helpers**
Edit `application/config/autoload.php`:
```php
$autoload['libraries'] = array('database', 'session');
$autoload['helper'] = array('url', 'file');
```

---

## **2. Database Setup**
Create a `users` table in your database:
```sql
CREATE TABLE users (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nama VARCHAR(100) NOT NULL,
    email VARCHAR(100) NOT NULL
);
```

---

## **3. Model (M_Excel.php)**
Create `application/models/M_Excel.php`:
```php
<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class M_Excel extends CI_Model {

    public function insert_batch($data) {
        $this->db->insert_batch('users', $data);
    }

    public function get_users() {
        return $this->db->get('users')->result_array();
    }
}
```

---

## **4. Controller (Excel.php)**
Create `application/controllers/Excel.php`:
```php
<?php
defined('BASEPATH') OR exit('No direct script access allowed');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

class Excel extends CI_Controller {

    public function __construct() {
        parent::__construct();
        $this->load->model('M_Excel');
        $this->load->database();
        $this->load->library('session');
        $this->load->helper('url');
    }

    public function index() {
        $this->load->view('upload_excel');
    }

    // Import data from Excel
    public function import() {
        $file_mimes = array('application/vnd.ms-excel', 'text/csv', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        if(isset($_FILES['file']['name']) && in_array($_FILES['file']['type'], $file_mimes)) {
            $file = $_FILES['file']['tmp_name'];

            $spreadsheet = IOFactory::load($file);
            $sheetData = $spreadsheet->getActiveSheet()->toArray();

            $data = [];
            foreach ($sheetData as $index => $row) {
                if ($index == 0) continue; // Skip header

                $data[] = [
                    'nama'  => $row[0],
                    'email' => $row[1],
                ];
            }

            if (!empty($data)) {
                $this->M_Excel->insert_batch($data);
            }

            $this->session->set_flashdata('message', 'Data successfully imported!');
            redirect('excel');
        } else {
            $this->session->set_flashdata('message', 'Invalid file format!');
            redirect('excel');
        }
    }

    // Export data to Excel
    public function export() {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // Set column headers
        $sheet->setCellValue('A1', 'Nama');
        $sheet->setCellValue('B1', 'Email');

        // Retrieve data from database
        $users = $this->M_Excel->get_users();
        $rowIndex = 2;

        foreach ($users as $user) {
            $sheet->setCellValue('A' . $rowIndex, $user['nama']);
            $sheet->setCellValue('B' . $rowIndex, $user['email']);
            $rowIndex++;
        }

        $writer = new Xlsx($spreadsheet);
        $filename = 'Data_User_' . date('YmdHis') . '.xlsx';

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="'. $filename .'"');
        header('Cache-Control: max-age=0');

        $writer->save('php://output');
    }
}
```

---

## **5. View (upload_excel.php)**
Create `application/views/upload_excel.php`:
```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Upload & Download Excel</title>
</head>
<body>

    <h2>Upload Excel File</h2>

    <?php if($this->session->flashdata('message')): ?>
        <p><?php echo $this->session->flashdata('message'); ?></p>
    <?php endif; ?>

    <form action="<?php echo base_url('excel/import'); ?>" method="post" enctype="multipart/form-data">
        <input type="file" name="file" required>
        <button type="submit">Upload</button>
    </form>

    <h2>Download Data as Excel</h2>
    <a href="<?php echo base_url('excel/export'); ?>"><button>Download Excel</button></a>

</body>
</html>
```

---

## **6. Usage**
### **6.1 Run the Project**
Start your local server:
```bash
php -S localhost:8000
```
Or use **Apache/Nginx**.

### **6.2 Open Browser and Navigate to:**
```
http://localhost/your_project/index.php/excel
```

### **6.3 Upload Excel File**
- Click "Upload"
- Ensure the file is in **.xls, .xlsx, or .csv** format.
- Data will be inserted into the **users** table.

### **6.4 Export Data to Excel**
- Click **"Download Excel"**.
- The system generates an **Excel file** containing all user data.

---

## **7. Troubleshooting**
| **Issue** | **Solution** |
|-----------|-------------|
| `base_url()` undefined | Load `url` helper in `autoload.php` |
| `session` not recognized | Load `session` library in `autoload.php` |
| `PhpSpreadsheet` class not found | Ensure `composer require phpoffice/phpspreadsheet` is executed |

---

## **8. Conclusion**
✅ Successfully integrated **PhpSpreadsheet** into **CodeIgniter 3**  
✅ Implemented **Import & Export Excel functionality**  
✅ Integrated with **MySQL database** using MVC  

🚀 Now, you can easily manage **Excel file uploads and downloads** in your CodeIgniter 3 project!

