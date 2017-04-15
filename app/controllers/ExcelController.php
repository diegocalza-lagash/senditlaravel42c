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
	public function show($id)
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
		$db = $m->formSendit2;
		$collection = $db->DataFormTest;
		$docSendit = $collection->findOne(['Entry.Id' => $id]);
		//echo "hola 1";
		//var_dump($docSendit);
		//echo $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['Sistema_bloqueo'];
		$StartTime = $docSendit['Entry']['StartTime'];
		$UserFirstName =$docSendit['Entry']['UserFirstName'];
		$UserLastName = $docSendit['Entry']['UserLastName'];
		$mantencion_equipos = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['mantencion_equipos'];
		$Trabajos = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['Trabajos'];
		$Sub_trabajos = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['Sub_trabajos'];
		$Sistema_bloqueo = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['Sistema_bloqueo'];
		$fecha_inicio_prog = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_inicio_prog'];
		$fecha_termino_prog = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_termino_prog'];
		$fecha_inicio_real = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_inicio_real'];
		$fecha_termino_real = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_termino_real'];
		$porcentaje_avance_fisico = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['porcentaje_avance_fisico'];
		$observaciones = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['observaciones'];
		//echo "hola".$docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['Sistema_bloqueo'];
		$objPHPExcel = new PHPExcel();
		// Leemos un archivo Excel 2007
		$objReader = PHPExcel_IOFactory::createReader('Excel2007');
		$objPHPExcel = $objReader->load("/var/www/senditlaravel42/public/reports/reporteRudelEmpty.xlsx");
		// Indicamos que se pare en la hoja uno del libro
		$objPHPExcel->setActiveSheetIndex(0);
		//Escribimos en la hoja en la celda B1
		$objPHPExcel->getActiveSheet()->SetCellValue('C11', $Sistema_bloqueo);
		$objPHPExcel->getActiveSheet()->SetCellValue('C14', $fecha_inicio_prog);
		$objPHPExcel->getActiveSheet()->SetCellValue('C15', $fecha_termino_prog);
		$objPHPExcel->getActiveSheet()->SetCellValue('F14', $fecha_inicio_real);
		$objPHPExcel->getActiveSheet()->SetCellValue('F15', $fecha_termino_real);
		$objPHPExcel->getActiveSheet()->SetCellValue('C19', $Trabajos);
		$objPHPExcel->getActiveSheet()->SetCellValue('C20', $Sub_trabajos);
		$objPHPExcel->getActiveSheet()->getStyle('C20')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
		$objPHPExcel->getActiveSheet()->SetCellValue('D20', $fecha_inicio_real);
		$objPHPExcel->getActiveSheet()->SetCellValue('E20', $fecha_termino_real);
		$objPHPExcel->getActiveSheet()->SetCellValue('F20', $porcentaje_avance_fisico."%");
		$objPHPExcel->getActiveSheet()->SetCellValue('C35', $observaciones);
		// Color rojo al texto
		/*$objPHPExcel->getActiveSheet()->getStyle('B2')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);
		// Texto alineado a la derecha
		$objPHPExcel->getActiveSheet()->getStyle('B2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
		// Damos un borde a la celda
		$objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
		$objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
		$objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);*/
		//Guardamos el archivo en formato Excel 2007
		//Si queremos trabajar con Excel 2003, basta cambiar el 'Excel2007' por 'Excel5' y el nombre del archivo de salida cambiar su formato por '.xls'

		//header('Content-type: application/vnd.ms-excel');excel 2003
		//PARA DESCARGAR EXCEL
		//RETIRAR ECHO O HACERLO EN UN ARCHIVO APARTE
		header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
		header("Cache-Control: max-age=0");
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		$objWriter->save("ReportOut.xlsx");
		$objWriter->save("php://output");
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

				$StartTime = $docSendit['Entry']['StartTime'];
				$UserFirstName =$docSendit['Entry']['UserFirstName'];
				$UserLastName = $docSendit['Entry']['UserLastName'];
				$mantencion_equipos = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['mantencion_equipos'];
				$Trabajos = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['Trabajos'];
				$Sub_trabajos = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['Sub_trabajos'];
				$Sistema_bloqueo = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['Sistema_bloqueo'];
				$fecha_inicio_prog = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_inicio_prog'];
				$fecha_termino_prog = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_termino_prog'];
				$fecha_inicio_real = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_inicio_real'];
				$fecha_termino_real = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_termino_real'];
				$porcentaje_avance_fisico = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['porcentaje_avance_fisico'];
				$observaciones = $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['observaciones'];

				$data = array(
					    	array($StartTime, 'data2'),
					    	array('data3', 'data4')
							);
				$sheet->with($data);

			});

		})->export('xls');*/
	}


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
