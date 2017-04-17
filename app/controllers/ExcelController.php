<?php

class ExcelController extends \BaseController {

	/**
	 * Display a listing of the resource.
	 *
	 * @return Response
	 */
	public function index()
	{
		//
	}


	/**
	 * Show the form for creating a new resource.
	 *
	 * @return Response
	 */
	public function create()
	{
		//
	}


	/**
	 * Store a newly created resource in storage.
	 *
	 * @return Response
	 */
	public function store()
	{
		//
	}


	/**
	 * Display the specified resource.
	 *
	 * @param  int  $id
	 * @return Response
	 */
	/*public function show($id)
	{
		//echo "id ".$id;
		// Camino a los include
		set_include_path(get_include_path() . PATH_SEPARATOR . '../Classes/');
		// PHPExcel
		//require_once 'Classes/PHPExcel.php';
		// PHPExcel_IOFactory
		//include 'PHPExcel/IOFactory.php';
		// Creamos un objeto PHPExcel
		$m = new MongoClient();
		$db = $m->SenditForm;
		$collWorks = $db->Works;
		$docWork = $collWorks->findOne(['Entry.Id' => $id]);//con id que viene de la view index
		//get field of works collec
		$StartTime = $docWork['Entry']['StartTime'];
		$UserFirstName =$docWork['Entry']['UserFirstName'];
		$UserLastName = $docWork['Entry']['UserLastName'];
		$Trabajos = $docWork['Entry']['AnswersJson']['ADD_WORK_PAGE']['WORK'];
		$Sub_trabajos = $docWork['Entry']['AnswersJson']['ADD_WORK_PAGE']['SUBWORK'];
		$Sistema_bloqueo = $docWork['Entry']['AnswersJson']['ADD_WORK_PAGE']['BLOCK_SYSTEM'];
		$fecha_inicio_prog = $docWork['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_PROGRAMMED'];
		$fecha_termino_prog = $docWork['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_PROGRAMMED'];
		$fecha_inicio_real = $docWork['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_REAL'];
		$fecha_termino_real = $docWork['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_REAL'];

		$porcentaje_avance_fisico = $docWork['Entry']['AnswersJson']['ADD_WORK_PAGE']['POOP'];
		$observaciones = $docWork['Entry']['AnswersJson']['ADD_WORK_PAGE']['OBSERVATIONS'];
		$s_t_day = $docWork['Entry']['AnswersJson']['ADD_WORK_PAGE']['S_TURN_DAY'];
		$s_t_night = $docWork['Entry']['AnswersJson']['ADD_WORK_PAGE']['S_TURN_NIGHT'];
		$i_p_day = $docWork['Entry']['AnswersJson']['ADD_WORK_PAGE']['I_P_TURN_DAY'];
		$i_p_night = $docWork['Entry']['AnswersJson']['ADD_WORK_PAGE']['I_P_TURN_NIGHT'];

		$objPHPExcel = new PHPExcel();
		// Leemos un archivo Excel 2007
		$objReader = PHPExcel_IOFactory::createReader('Excel2007');

//		$objPHPExcel = $objReader->load("/var/www/senditlaravel42/public/reports/reporteRudelEmpty.xlsx");

		try {
			$objPHPExcel = $objReader->load("public/reports/reporteRudelEmpty.xlsx");
		} catch (Exception $e) {
			$objPHPExcel = $objReader->load("/var/www/senditlaravel42/public/reports/reporteRudelEmpty.xlsx");
			//echo "se capturo excep";
		}


		// Indicamos que se pare en la hoja uno del libro
		$objPHPExcel->setActiveSheetIndex(0);
		$objPHPExcel->getActiveSheet()->SetCellValue('C11', $Sistema_bloqueo);
		$objPHPExcel->getActiveSheet()->SetCellValue('C14', $fecha_inicio_prog);
		$objPHPExcel->getActiveSheet()->SetCellValue('C15', $fecha_termino_prog);
		//$objPHPExcel->getActiveSheet()->SetCellValue('F14', $fecha_inicio_real);
		//$objPHPExcel->getActiveSheet()->SetCellValue('F15', $fecha_termino_real);
	$objPHPExcel->getActiveSheet()->SetCellValue('C19', $Trabajos);
		$objPHPExcel->getActiveSheet()->SetCellValue('C20', $Sub_trabajos);
		$objPHPExcel->getActiveSheet()->getStyle('C20')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
		$fecha_inicio_real = new DateTime($fecha_inicio_real);
		$objPHPExcel->getActiveSheet()->SetCellValue('D20', date_format($fecha_inicio_real, 'd/m/Y H:i:s'));
		$fecha_termino_real = new DateTime($fecha_termino_real);
		$objPHPExcel->getActiveSheet()->SetCellValue('E20', date_format($fecha_termino_real, 'd/m/Y H:i:s'));
		$objPHPExcel->getActiveSheet()->SetCellValue('F20', $porcentaje_avance_fisico."%");
		$objPHPExcel->getActiveSheet()->SetCellValue('C35', $observaciones);
		$objPHPExcel->getActiveSheet()->SetCellValue('F9', $s_t_day);
		$objPHPExcel->getActiveSheet()->SetCellValue('F10', $s_t_night);
		$objPHPExcel->getActiveSheet()->SetCellValue('F11', $i_p_day);
		$objPHPExcel->getActiveSheet()->SetCellValue('F12', $i_p_night);

		$db = $m->SenditForm;
		$collSubWorks = $db->SubWorks;
		$docsSubworks = $collSubWorks->find(['Entry.Id' => $id]);
		//echo "hola 1";
		if ($docsSubworks->count() > 0) {
			$row = 21;
			foreach ($docsSubworks as $subw) {
				$subw = $subw['Entry']['AnswersJson']['ADD_WORK_PAGE']['SUBWORK'];
				//$poop = $subw['Entry']['AnswersJson']['ADD_WORK_PAGE']['POOP'];
				//var_dump($subw);
				//echo "hola";


				$objPHPExcel->getActiveSheet()->SetCellValue('C'.(string)($row), $subw);
				$objPHPExcel->getActiveSheet()->getStyle('C'.(string)($row))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
				$row ++;
			}
			$row=21;
			foreach ($docsSubworks as $date_start_r) {
				$date_start_r = $date_start_r['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_REAL'];
				$date_start_r = new DateTime($date_start_r);
				$objPHPExcel->getActiveSheet()->SetCellValue('D'.(string)($row), date_format($date_start_r, 'd/m/Y H:i:s'));
				$row++;
			}
			$row=21;
			foreach ($docsSubworks as $date_end_r) {
				$date_end_r = $date_end_r['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_REAL'];
				$date_end_r = new DateTime($date_end_r);
				$objPHPExcel->getActiveSheet()->SetCellValue('E'.(string)($row), date_format($date_end_r, 'd/m/Y H:i:s'));
				$row++;
			}
			$row=21;
			foreach ($docsSubworks as $poop) {
				$poop = $poop['Entry']['AnswersJson']['ADD_WORK_PAGE']['POOP'];
				$objPHPExcel->getActiveSheet()->SetCellValue('F'.(string)($row), $poop."%");
				$row++;
			}
		}else{
			//echo "no hay doc en Subworks collec con Id: ".$id;
		}
		header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
		header("Cache-Control: max-age=0");
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		$objWriter->save("ReportOut.xlsx");
		$objWriter->save("php://output");

		// Color rojo al texto
		/*$objPHPExcel->getActiveSheet()->getStyle('B2')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);
		// Texto alineado a la derecha
		$objPHPExcel->getActiveSheet()->getStyle('B2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
		// Damos un borde a la celda
		$objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
		$objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
		$objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
		*/
		//Guardamos el archivo en formato Excel 2007
		//Si queremos trabajar con Excel 2003, basta cambiar el 'Excel2007' por 'Excel5' y el nombre del archivo de salida cambiar su formato por '.xls'
		//header('Content-type: application/vnd.ms-excel');excel 2003
		//PARA DESCARGAR EXCEL
		//RETIRAR ECHO O HACERLO EN UN ARCHIVO APARTE

		//
		//echo "hola excel";
		//global $id;
		//$GLOBALS[$id];


		/*Excel::create("report3",function($excel){

			$excel->sheet('kalza',function($sheet){
				//$UserFirstName =$docSendit['Entry']['UserFirstName'];
				/*$data=[];
				array_push($data, array(
							array('data1', 'data2'),
					    	array('data3', 'data4')));*/
				/*$data = array(
					    	array('data1', 'data2'),
					    	array('data3', 'data4')
							);/*
				$m = new MongoDB\Client();
				$db = $m->formSendit2;
				$collection = $db->DataFormTest;
				$docSendit = $collection->findOne(['Entry.Id' => $id]);

			});

		})->export('xls');
	}*/


	/**
	 * Show the form for editing the specified resource.
	 *
	 * @param  int  $id
	 * @return Response
	 */
	public function edit($id)
	{
		//
	}


	/**
	 * Update the specified resource in storage.
	 *
	 * @param  int  $id
	 * @return Response
	 */
	public function update($id)
	{
		//
	}


	/**
	 * Remove the specified resource from storage.
	 *
	 * @param  int  $id
	 * @return Response
	 */
	public function destroy($id)
	{
		//
	}


}
