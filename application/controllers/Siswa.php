<?php
defined('BASEPATH') or exit('No direct script access allowed');

class Siswa extends CI_Controller
{
	private $filename = "import_data";

	public function __construct()
	{
		parent::__construct();

		$this->load->model('Siswa_model');
	}

	public function index()
	{
		$data['siswa'] = $this->Siswa_model->view();
		$this->load->view('view', $data);
	}

	public function form()
	{
		$data = array();

		if (isset($_POST['preview'])) {
			$upload = $this->Siswa_model->upload_file($this->filename);

			if ($upload['result'] == "success") {
				include APPPATH . 'third_party/PHPExcel/PHPExcel.php';

				$excelreader = new PHPExcel_Reader_Excel2007();
				$loadexcel = $excelreader->load('excel/' . $this->filename . '.xlsx');
				$sheet = $loadexcel->getActiveSheet()->toArray(null, true, true, true);

				$data['sheet'] = $sheet;
			} else {
				$data['upload_error'] = $upload['error'];
			}
		}

		$this->load->view('form', $data);
	}

	public function import()
	{
		include APPPATH . 'third_party/PHPExcel/PHPExcel.php';

		$excelreader = new PHPExcel_Reader_Excel2007();
		$loadexcel = $excelreader->load('excel/' . $this->filename . '.xlsx');

		$sheet = $loadexcel->getActiveSheet()->toArray(null, true, true, true);

		$data = array();

		$numrow = 1;
		foreach ($sheet as $row) {
			if ($numrow > 1) {
				array_push($data, array(
					'nis' => $row['A'], // Insert data nis dari kolom A di excel
					'nama' => $row['B'], // Insert data nis dari kolom B di excel
					'jenis_kelamin' => $row['C'], // Insert data nis dari kolom C di excel
					'alamat' => $row['D'], // Insert data nis dari kolom D di excel
				));
			}

			$numrow++; //tambah 1 setiap kali looping
		}

		$this->Siswa_model->insert_multiple($data);

		redirect("Siswa");
	}
}
