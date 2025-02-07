<?php
defined('BASEPATH') or exit('No direct script access allowed');

class M_Excel extends CI_Model
{

    // Simpan data hasil impor ke database
    public function insert_batch($data)
    {
        $this->db->insert_batch('users', $data);
    }

    // Ambil semua data untuk ekspor
    public function get_users()
    {
        return $this->db->get('users')->result_array();
    }
}
