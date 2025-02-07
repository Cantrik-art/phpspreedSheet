<?php
defined('BASEPATH') or exit('No direct script access allowed');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

class Excel extends CI_Controller
{

    public function __construct()
    {
        parent::__construct();
        $this->load->model('M_Excel');
        $this->load->database();
        $this->load->library('session'); // Tambahkan ini!
        $this->load->helper('url'); // Tambahkan ini
    }

    // ==================== IMPOR DATA DARI EXCEL KE DATABASE ====================
    public function import()
    {
        $file_mimes = array('application/vnd.ms-excel', 'text/csv', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        if (isset($_FILES['file']['name']) && in_array($_FILES['file']['type'], $file_mimes)) {
            $file = $_FILES['file']['tmp_name'];

            $spreadsheet = IOFactory::load($file);
            $sheetData = $spreadsheet->getActiveSheet()->toArray();

            $data = [];
            foreach ($sheetData as $index => $row) {
                if ($index == 0) continue; // Lewati baris header

                $data[] = [
                    'nama'  => $row[0],
                    'email' => $row[1],
                    'nik' => $row[2]
                ];
            }

            if (!empty($data)) {
                $this->M_Excel->insert_batch($data);
            }

            $this->session->set_flashdata('message', 'Data berhasil diimpor!');
            redirect('excel');
        } else {
            $this->session->set_flashdata('message', 'Format file tidak valid!');
            redirect('excel');
        }
    }

    // ==================== EKSPOR DATA KE FILE EXCEL ====================
    public function export()
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // Header kolom
        $sheet->setCellValue('A1', 'Nama');
        $sheet->setCellValue('B1', 'Email');
        $sheet->setCellValue('C1', 'Nik');

        // Ambil data dari database
        $users = $this->M_Excel->get_users();
        $rowIndex = 2;

        foreach ($users as $user) {
            $sheet->setCellValue('A' . $rowIndex, $user['nama']);
            $sheet->setCellValue('B' . $rowIndex, $user['email']);
            $sheet->setCellValue('C' . $rowIndex, $user['nik']);
            $rowIndex++;
        }

        $writer = new Xlsx($spreadsheet);
        $filename = 'Data_User_' . date('YmdHis') . '.xlsx';

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header('Cache-Control: max-age=0');

        $writer->save('php://output');
    }

    // ==================== TAMPILKAN FORM UPLOAD DI VIEW ====================
    public function index()
    {
        $this->load->view('upload_excel');
    }
}
