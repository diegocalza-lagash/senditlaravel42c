<?php
error_reporting(E_ALL);
//require "/vendor/autoload.php";
class Console
{
    /**
     * @param string $name Nombre único para poder ejecutar esto varias veces en el mismo documento
     * @param mixed $var Una variable cadena, objeto, matriz o lo que sea
     * @param string $type (debug|info|warn|error)
     * @return html
     */
    public static function log($name, $var, $type='debug')
    {
        $name = preg_replace('/[^A-Z|0-9]/i', '_', $name);
        $types = array('debug', 'info', 'warn', 'error');
        if ( ! in_array($type, $types) ) $type = 'debug';
        $s = '<script>' . PHP_EOL;
            if ( is_object($var) or is_array($var) )
            {
                $object = json_encode($var);
                $object = str_replace("'", "\'", $object);
                $s .= "var object$name = '$object';" . PHP_EOL;
                $s .= "var val$name = eval('('+object$name+')');" . PHP_EOL;
                $s .= "console.$type(val$name);" . PHP_EOL;
            }
            else
            {
                $var = str_replace('"', '\\"', $var);
                $s .= "console.$type($var);" . PHP_EOL;
            }
        $s .= '</script>' . PHP_EOL;
        return $s;
    }
}


class DataSendController extends \BaseController {

	/**
	 * Display a listing of the resource.
	 *
	 * @return Response
	 */
	public function getIndex()
	{
		//echo "hola";
		return View::make('DataSend.index');
		//return View::make('DataSend.report',array("docRepor" => $docRepor));
	}

	public function showWorks()
	{
				/*echo "hola";
				$m = new MongoClient();
				$db = $m->SenditForm;
				$collWorks = $db->Works;
				$docsWorks = $collWorks->find();
				//$docsWorks = Work::all();

				return View::make('listworks', array('dataform' => $docsWorks));
		//return View::make('dataSends.index')->with('dataform', $docsWorks);
		//return Redirect::to('dataform');*/

	}
	public function report(){
		//echo "hola";
		if (isset($_GET["equi"]))
            {
				$equi = htmlspecialchars(Input::get("equi"));
				$loc = htmlspecialchars(Input::get("loc"));
				$iden = htmlspecialchars(Input::get("iden"));
				$dep = htmlspecialchars(Input::get("dep"));
				$m = new MongoClient();//obsoleta desde mongo 1.0.0
				$db = $m->SenditForm;
				$collRepor = $db->Repor;
				$docRepor = $collRepor->find([
					'EQUIPMENT.EQUIPMENT_NAME' => $equi,
					'EQUIPMENT.LOCALIZATION_EQUIPMENT.LOCALIZATION_NAME' => $loc,
					'EQUIPMENT.IDENTIFICATION_EQUIPMENT.IDENTIFICATION_NAME' => $iden,
					'EQUIPMENT.DATE_END_PROGRAMMED' => $dep
					]);
				//return View::make('DataSend.report',array("docRepor" => $docRepor));
				if (!$docRepor -> count()) {
					//echo "Sin Trabajos";
				}else{
					//$docReporArray = iterator_to_array($docRepor,false);

					//$docsample = array("work" => "w2", "subw"=>"sw2");
					$m = new MongoClient();
					$db = $m->SenditForm;
					$collwf = $db->works_filter;

					foreach ($docRepor as  $v) {

						$docwork = $collwf->insert(array(
							"work" => $v["EQUIPMENT"]["WORK"]["WORK_NAME"],
							"subwork" => $v["EQUIPMENT"]["WORK"]["SUBWORK"]["SUBWORK_NAME"],
							"work_nuevo" => $v['EQUIPMENT']['WORK']['WORK_NUEVO'],
							"dsr" => $v['EQUIPMENT']['WORK']['SUBWORK']['DATE_START_REAL'],
							"der" => $v['EQUIPMENT']['WORK']['SUBWORK']['DATE_END_REAL'],
							"poop" => $v['EQUIPMENT']['WORK']['SUBWORK']['POOP'],
							"obs" => $v['EQUIPMENT']['WORK']['SUBWORK']['OBSERVATIONS']

							));

					}

					$docwork = $collwf->insert($docRepor);

					$keys = array("work" => 1);
					$initial = array("subworks" => array());
					$reduce = "function(obj, prev){prev.subworks.push(obj.subwork,obj.dsr,obj.der,obj.poop,obj.obs)}";
					$g = $collwf->group($keys, $initial,$reduce);
					$collwf->drop();
					$collwf = $db->works_filter;
					$docwork = $collwf->insert($g);
					$objPHPExcel = new PHPExcel();
					$objReader = PHPExcel_IOFactory::createReader('Excel2007');
					try {
						$objPHPExcel = $objReader->load("public/reports/reporteRudel.xlsx");
					} catch (Exception $e) {
						$objPHPExcel = $objReader->load("/var/www/senditlaravel42/public/reports/reporteRudel.xlsx");
					}

					$objWorksheet= $objPHPExcel->setActiveSheetIndex(0);
					//echo count($g['retval']);
					//si es un trabajo y un subtrabajo
					if (count($g['retval']) == 2 && count($g['retval'][1]['subworks']) == 5) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						//var_dump(count($g['retval'][1]['subworks']));
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//si es 1 trabajo y 2 subtrabajo
					if (count($g['retval']) == 2 && count($g['retval'][1]['subworks']) == 10) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12 = $der12->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");

					}
					//si es un w y 3 subw
					if (count($g['retval']) == 2 && count($g['retval'][1]['subworks']) == 15) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12 = $der12->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13 = $der13->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");

					}
					// si es 1 w y 4 subw
					if (count($g['retval']) == 2 && count($g['retval'][1]['subworks']) == 20) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12 = $der12->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13 = $der13->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14 = $der14->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");

					}
					//si es 1 w y 5 subw
					if (count($g['retval']) == 2 && count($g['retval'][1]['subworks']) == 25) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12 = $der12->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13 = $der13->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14 = $der14->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$subwork15 = $g['retval'][1]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr15 = $dsr15->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][1]['subworks'][22]);
						$der15 = $der15->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][1]['subworks'][23];
						$obs15 = $g['retval'][1]['subworks'][24];
						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop15."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");

					}
					//si es un 1 w y 6 subw
					if (count($g['retval']) == 2 && count($g['retval'][1]['subworks']) == 30) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12 = $der12->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13 = $der13->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14 = $der14->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$subwork15 = $g['retval'][1]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr15 = $dsr15->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][1]['subworks'][22]);
						$der15 = $der15->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][1]['subworks'][23];
						$obs15 = $g['retval'][1]['subworks'][24];
						$subwork16 = $g['retval'][1]['subworks'][25];
						$dsr16 = new DateTime($g['retval'][1]['subworks'][26]);
						$dsr16 = $dsr16->format('d-m-Y, g:i a');
						$der16 = new DateTime($g['retval'][1]['subworks'][27]);
						$der16 = $der16->format('d-m-Y, g:i a');
						$poop16 = $g['retval'][1]['subworks'][28];
						$obs16 = $g['retval'][1]['subworks'][29];
						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop15."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop16."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");

					}
					//ASCENDENTE
					//si son 2 w y un subwork ambos
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks']) == 5 && count($g['retval'][2]['subworks']) == 5) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('D22', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop21."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					// 2 w el 1er w con 1 subwork y el 2do w con 2subw en orden
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 5 && count($g['retval'][2]['subworks']) == 10) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('D22', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop22."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 w el 1erw con 1 subwork y el 2do w con 3subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks']) == 5 && count($g['retval'][2]['subworks']) == 15) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('D22', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop23."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					// 2 w el 1ero con 1 subwork y el 2do w con 4subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks']) == 5 && count($g['retval'][2]['subworks']) == 20) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24 = $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('D22', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop24."%");
						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					// 2 w el 1ero w con 1 subwork y el 2do w con 5subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks']) == 5 && count($g['retval'][2]['subworks']) == 25) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24 = $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];
						$subwork25 = $g['retval'][2]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][2]['subworks'][21]);
						$dsr25 = $dsr25->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][2]['subworks'][22]);
						$der25 = $der25->format('d-m-Y, g:i a'); ;
						$poop25 = $g['retval'][2]['subworks'][23];
						$obs25 = $g['retval'][2]['subworks'][24];
						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('D22', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop24."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop25."%");
						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					// 2 w el 1ero con 1 subwork y el 2do w con 6subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks']) == 5 && count($g['retval'][2]['subworks']) == 30) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24 = $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];
						$subwork25 = $g['retval'][2]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][2]['subworks'][21]);
						$dsr25 = $dsr25->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][2]['subworks'][22]);
						$der25 = $der25->format('d-m-Y, g:i a'); ;
						$poop25 = $g['retval'][2]['subworks'][23];
						$obs25 = $g['retval'][2]['subworks'][24];
						$subwork26 = $g['retval'][2]['subworks'][21];
						$dsr26 = new DateTime($g['retval'][2]['subworks'][22]);
						$dsr26 = $dsr26->format('d-m-Y, g:i a');
						$der26 = new DateTime($g['retval'][2]['subworks'][23]);
						$der26 = $der26->format('d-m-Y, g:i a'); ;
						$poop26 = $g['retval'][2]['subworks'][24];
						$obs26 = $g['retval'][2]['subworks'][25];
						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('D22', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop24."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop25."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop26."%");
						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2W 1ero 2SUB
					//2 W, 1ero con 2 subw y el 2do con 2 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 10 && count($g['retval'][2]['subworks']) == 10) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D23', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop22."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 2 subw y el 2do con 3 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 10 && count($g['retval'][2]['subworks']) == 15) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D23', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop23."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 2 subw y el 2do con 4 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 10 && count($g['retval'][2]['subworks']) == 20) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D23', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop24."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 2 subw y el 2do con 5 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 10 && count($g['retval'][2]['subworks']) == 25) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];
						$subwork25 = $g['retval'][2]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][2]['subworks'][21]);
						$dsr25 = $dsr25->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][2]['subworks'][22]);
						$der25= $der25->format('d-m-Y, g:i a'); ;
						$poop25 = $g['retval'][2]['subworks'][23];
						$obs25 = $g['retval'][2]['subworks'][24];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D23', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop24."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop25."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 2 subw y el 2do con 6 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 10 && count($g['retval'][2]['subworks']) == 30) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];
						$subwork25 = $g['retval'][2]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][2]['subworks'][21]);
						$dsr25 = $dsr25->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][2]['subworks'][22]);
						$der25= $der25->format('d-m-Y, g:i a'); ;
						$poop25 = $g['retval'][2]['subworks'][23];
						$obs25 = $g['retval'][2]['subworks'][24];
						$subwork26 = $g['retval'][2]['subworks'][25];
						$dsr26 = new DateTime($g['retval'][2]['subworks'][26]);
						$dsr26 = $dsr26->format('d-m-Y, g:i a');
						$der26 = new DateTime($g['retval'][2]['subworks'][27]);
						$der26= $der26->format('d-m-Y, g:i a'); ;
						$poop26 = $g['retval'][2]['subworks'][28];
						$obs26 = $g['retval'][2]['subworks'][29];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D23', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop24."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop25."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E29', $subwork26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB29', $dsr26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH29', $der26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN29', $poop26."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2W 1ero 3SUB
					//2 W, 1ero con 3 subw y el 2do con 3 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 15 && count($g['retval'][2]['subworks']) == 15) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D24', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop23."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 3 subw y el 2do con 4 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 15 && count($g['retval'][2]['subworks']) == 20) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D24', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop24."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 3 subw y el 2do con 5 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 15 && count($g['retval'][2]['subworks']) == 25) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];
						$subwork25 = $g['retval'][2]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][2]['subworks'][21]);
						$dsr25 = $dsr25->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][2]['subworks'][22]);
						$der25= $der25->format('d-m-Y, g:i a'); ;
						$poop25 = $g['retval'][2]['subworks'][23];
						$obs25 = $g['retval'][2]['subworks'][24];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D24', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop24."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E29', $subwork25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB29', $dsr25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH29', $der25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN29', $poop25."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 3 subw y el 2do con 6 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 15 && count($g['retval'][2]['subworks']) == 30) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];
						$subwork25 = $g['retval'][2]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][2]['subworks'][21]);
						$dsr25 = $dsr25->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][2]['subworks'][22]);
						$der25= $der25->format('d-m-Y, g:i a'); ;
						$poop25 = $g['retval'][2]['subworks'][23];
						$obs25 = $g['retval'][2]['subworks'][24];
						$subwork26 = $g['retval'][2]['subworks'][25];
						$dsr26 = new DateTime($g['retval'][2]['subworks'][26]);
						$dsr26 = $dsr26->format('d-m-Y, g:i a');
						$der26 = new DateTime($g['retval'][2]['subworks'][27]);
						$der26= $der26->format('d-m-Y, g:i a'); ;
						$poop26 = $g['retval'][2]['subworks'][28];
						$obs26 = $g['retval'][2]['subworks'][29];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D24', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop24."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E29', $subwork25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB29', $dsr25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH29', $der25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN29', $poop25."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E30', $subwork26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB30', $dsr26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH30', $der26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN30', $poop26."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2W 1ero 4Sub
					//2 W, 1ero con 4 subw y el 2do con 4 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 20 && count($g['retval'][2]['subworks']) == 20) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D25', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E29', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB29', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH29', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN29', $poop24."%");


						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 4 subw y el 2do con 5 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 20 && count($g['retval'][2]['subworks']) == 25) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];
						$subwork25 = $g['retval'][2]['subworks'][15];
						$dsr25 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr25 = $dsr25->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][2]['subworks'][17]);
						$der25= $der25->format('d-m-Y, g:i a'); ;
						$poop25 = $g['retval'][2]['subworks'][18];
						$obs25 = $g['retval'][2]['subworks'][19];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D25', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E29', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB29', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH29', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN29', $poop24."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E30', $subwork25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB30', $dsr25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH30', $der25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN30', $poop25."%");


						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 4 subw y el 2do con 6 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 20 && count($g['retval'][2]['subworks']) == 30) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];
						$subwork25 = $g['retval'][2]['subworks'][15];
						$dsr25 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr25 = $dsr25->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][2]['subworks'][17]);
						$der25= $der25->format('d-m-Y, g:i a'); ;
						$poop25 = $g['retval'][2]['subworks'][18];
						$obs25 = $g['retval'][2]['subworks'][19];
						$subwork26 = $g['retval'][2]['subworks'][20];
						$dsr26 = new DateTime($g['retval'][2]['subworks'][21]);
						$dsr26 = $dsr26->format('d-m-Y, g:i a');
						$der26 = new DateTime($g['retval'][2]['subworks'][22]);
						$der26= $der26->format('d-m-Y, g:i a'); ;
						$poop26 = $g['retval'][2]['subworks'][23];
						$obs26 = $g['retval'][2]['subworks'][24];


						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D25', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E29', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB29', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH29', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN29', $poop24."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E30', $subwork25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB30', $dsr25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH30', $der25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN30', $poop25."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E31', $subwork26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB31', $dsr26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH31', $der26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN31', $poop26."%");


						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 5 subw y el 2do con 5 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 25 && count($g['retval'][2]['subworks']) == 25) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$subwork15 = $g['retval'][1]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr15 = $dsr15->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][1]['subworks'][22]);
						$der15= $der15->format('d-m-Y, g:i a'); ;
						$poop15 = $g['retval'][1]['subworks'][23];
						$obs15 = $g['retval'][1]['subworks'][24];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];
						$subwork25 = $g['retval'][2]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][2]['subworks'][21]);
						$dsr25 = $dsr25->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][2]['subworks'][22]);
						$der25= $der25->format('d-m-Y, g:i a'); ;
						$poop25 = $g['retval'][2]['subworks'][23];
						$obs25 = $g['retval'][2]['subworks'][24];


						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop15."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D26', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E29', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB29', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH29', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN29', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E30', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB30', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH30', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN30', $poop24."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E31', $subwork25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB31', $dsr25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH31', $der25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN31', $poop25."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 5 subw y el 2do con 6 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 25 && count($g['retval'][2]['subworks']) == 30) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$subwork15 = $g['retval'][1]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr15 = $dsr15->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][1]['subworks'][22]);
						$der15= $der15->format('d-m-Y, g:i a'); ;
						$poop15 = $g['retval'][1]['subworks'][23];
						$obs15 = $g['retval'][1]['subworks'][24];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];
						$subwork25 = $g['retval'][2]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][2]['subworks'][21]);
						$dsr25 = $dsr25->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][2]['subworks'][22]);
						$der25= $der25->format('d-m-Y, g:i a'); ;
						$poop25 = $g['retval'][2]['subworks'][23];
						$obs25 = $g['retval'][2]['subworks'][24];
						$subwork26 = $g['retval'][2]['subworks'][25];
						$dsr26 = new DateTime($g['retval'][2]['subworks'][26]);
						$dsr26 = $dsr26->format('d-m-Y, g:i a');
						$der26 = new DateTime($g['retval'][2]['subworks'][27]);
						$der26= $der26->format('d-m-Y, g:i a'); ;
						$poop26 = $g['retval'][2]['subworks'][28];
						$obs26 = $g['retval'][2]['subworks'][29];


						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop15."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D26', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E29', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB29', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH29', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN29', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E30', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB30', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH30', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN30', $poop24."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E31', $subwork25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB31', $dsr25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH31', $der25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN31', $poop25."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E32', $subwork26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB32', $dsr26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH32', $der26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN32', $poop26."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 6 subw y el 2do con 6 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 30 && count($g['retval'][2]['subworks']) == 30) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$subwork15 = $g['retval'][1]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr15 = $dsr15->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][1]['subworks'][22]);
						$der15= $der15->format('d-m-Y, g:i a'); ;
						$poop15 = $g['retval'][1]['subworks'][23];
						$obs15 = $g['retval'][1]['subworks'][24];
						$subwork16 = $g['retval'][1]['subworks'][25];
						$dsr16 = new DateTime($g['retval'][1]['subworks'][26]);
						$dsr16 = $dsr16->format('d-m-Y, g:i a');
						$der16 = new DateTime($g['retval'][1]['subworks'][27]);
						$der16= $der16->format('d-m-Y, g:i a'); ;
						$poop16 = $g['retval'][1]['subworks'][28];
						$obs16 = $g['retval'][1]['subworks'][29];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];
						$subwork25 = $g['retval'][2]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][2]['subworks'][21]);
						$dsr25 = $dsr25->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][2]['subworks'][22]);
						$der25= $der25->format('d-m-Y, g:i a'); ;
						$poop25 = $g['retval'][2]['subworks'][23];
						$obs25 = $g['retval'][2]['subworks'][24];
						$subwork26 = $g['retval'][2]['subworks'][25];
						$dsr26 = new DateTime($g['retval'][2]['subworks'][26]);
						$dsr26 = $dsr26->format('d-m-Y, g:i a');
						$der26 = new DateTime($g['retval'][2]['subworks'][27]);
						$der26= $der26->format('d-m-Y, g:i a'); ;
						$poop26 = $g['retval'][2]['subworks'][28];
						$obs26 = $g['retval'][2]['subworks'][29];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop15."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop16."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D27', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E29', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB29', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH29', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN29', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E30', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB30', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH30', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN30', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E31', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB31', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH31', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN31', $poop24."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E32', $subwork25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB32', $dsr25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH32', $der25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN32', $poop25."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E33', $subwork26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB33', $dsr26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH33', $der26);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN33', $poop26."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}


					//DESCENDENTE
					//2 W, PRIMERO CON 2SUBW EL 2DO CON 1 SUBW
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 10 && count($g['retval'][2]['subworks']) == 5) {

						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('D23', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop21."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2W, 1ERO CON 3SUB EL 2DO CON 1 SUBW
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 15 && count($g['retval'][2]['subworks']) == 5) {

						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");


						$objPHPExcel->getActiveSheet()->SetCellValue('D24', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop21."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2W, 1ERO CON 4SUB EL 2DO CON 1 SUBW
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 20 && count($g['retval'][2]['subworks']) == 5) {

						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D25', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop21."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2W, 1ERO CON 5SUB EL 2DO CON 1 SUBW
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 25 && count($g['retval'][2]['subworks']) == 5) {

						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$subwork15 = $g['retval'][1]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr15 = $dsr15->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][1]['subworks'][22]);
						$der15= $der15->format('d-m-Y, g:i a'); ;
						$poop15 = $g['retval'][1]['subworks'][23];
						$obs15 = $g['retval'][1]['subworks'][24];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop15."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D26', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop21."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2W, 1ERO CON 6SUB EL 2DO CON 1 SUBW
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 30 && count($g['retval'][2]['subworks']) == 5) {

						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$subwork15 = $g['retval'][1]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr15 = $dsr15->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][1]['subworks'][22]);
						$der15= $der15->format('d-m-Y, g:i a'); ;
						$poop15 = $g['retval'][1]['subworks'][23];
						$obs15 = $g['retval'][1]['subworks'][24];
						$subwork16 = $g['retval'][1]['subworks'][25];
						$dsr16 = new DateTime($g['retval'][1]['subworks'][26]);
						$dsr16 = $dsr16->format('d-m-Y, g:i a');
						$der16 = new DateTime($g['retval'][1]['subworks'][27]);
						$der16= $der16->format('d-m-Y, g:i a'); ;
						$poop16 = $g['retval'][1]['subworks'][28];
						$obs16 = $g['retval'][1]['subworks'][29];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop15."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop16."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D27', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop21."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 3 subw y el 2do con 2 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 15 && count($g['retval'][2]['subworks']) == 10) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D24', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop22."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 4 subw y el 2do con 2 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 20 && count($g['retval'][2]['subworks']) == 10) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");


						$objPHPExcel->getActiveSheet()->SetCellValue('D25', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop22."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 5 subw y el 2do con 2 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 25 && count($g['retval'][2]['subworks']) == 10) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$subwork15 = $g['retval'][1]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr15 = $dsr15->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][1]['subworks'][22]);
						$der15= $der15->format('d-m-Y, g:i a'); ;
						$poop15 = $g['retval'][1]['subworks'][23];
						$obs15 = $g['retval'][1]['subworks'][24];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop15."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D26', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop22."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 6 subw y el 2do con 2 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 30 && count($g['retval'][2]['subworks']) == 10) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$subwork15 = $g['retval'][1]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr15 = $dsr15->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][1]['subworks'][22]);
						$der15= $der15->format('d-m-Y, g:i a'); ;
						$poop15 = $g['retval'][1]['subworks'][23];
						$obs15 = $g['retval'][1]['subworks'][24];
						$subwork16 = $g['retval'][1]['subworks'][25];
						$dsr16 = new DateTime($g['retval'][1]['subworks'][26]);
						$dsr16 = $dsr16->format('d-m-Y, g:i a');
						$der16 = new DateTime($g['retval'][1]['subworks'][27]);
						$der16= $der16->format('d-m-Y, g:i a'); ;
						$poop16 = $g['retval'][1]['subworks'][28];
						$obs16 = $g['retval'][1]['subworks'][29];


						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop15."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop16."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D27', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E29', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB29', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH29', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN29', $poop22."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 4 subw y el 2do con 3 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 20 && count($g['retval'][2]['subworks']) == 15) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D25', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop23."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 5 subw y el 2do con 3 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 25 && count($g['retval'][2]['subworks']) == 15) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$subwork15 = $g['retval'][1]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr15 = $dsr15->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][1]['subworks'][22]);
						$der15= $der15->format('d-m-Y, g:i a'); ;
						$poop15 = $g['retval'][1]['subworks'][23];
						$obs15 = $g['retval'][1]['subworks'][24];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop15."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D26', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E29', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB29', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH29', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN29', $poop23."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 6 subw y el 2do con 3 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 30 && count($g['retval'][2]['subworks']) == 15) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$subwork15 = $g['retval'][1]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr15 = $dsr15->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][1]['subworks'][22]);
						$der15= $der15->format('d-m-Y, g:i a'); ;
						$poop15 = $g['retval'][1]['subworks'][23];
						$obs15 = $g['retval'][1]['subworks'][24];
						$subwork16 = $g['retval'][1]['subworks'][25];
						$dsr16 = new DateTime($g['retval'][1]['subworks'][26]);
						$dsr16 = $dsr16->format('d-m-Y, g:i a');
						$der16 = new DateTime($g['retval'][1]['subworks'][27]);
						$der16= $der16->format('d-m-Y, g:i a'); ;
						$poop16 = $g['retval'][1]['subworks'][28];
						$obs16 = $g['retval'][1]['subworks'][29];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop15."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop16."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D27', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E29', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB29', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH29', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN29', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E30', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB30', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH30', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN30', $poop23."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 5 subw y el 2do con 4 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 25 && count($g['retval'][2]['subworks']) == 20) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$subwork15 = $g['retval'][1]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr15 = $dsr15->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][1]['subworks'][22]);
						$der15= $der15->format('d-m-Y, g:i a'); ;
						$poop15 = $g['retval'][1]['subworks'][23];
						$obs15 = $g['retval'][1]['subworks'][24];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop15."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D26', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E27', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB27', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH27', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN27', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E29', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB29', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH29', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN29', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E30', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB30', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH30', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN30', $poop24."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 6 subw y el 2do con 4 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 30 && count($g['retval'][2]['subworks']) == 20) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$subwork15 = $g['retval'][1]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr15 = $dsr15->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][1]['subworks'][22]);
						$der15= $der15->format('d-m-Y, g:i a'); ;
						$poop15 = $g['retval'][1]['subworks'][23];
						$obs15 = $g['retval'][1]['subworks'][24];
						$subwork16 = $g['retval'][1]['subworks'][25];
						$dsr16 = new DateTime($g['retval'][1]['subworks'][26]);
						$dsr16 = $dsr16->format('d-m-Y, g:i a');
						$der16 = new DateTime($g['retval'][1]['subworks'][27]);
						$der16= $der16->format('d-m-Y, g:i a'); ;
						$poop16 = $g['retval'][1]['subworks'][28];
						$obs16 = $g['retval'][1]['subworks'][29];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];


						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop15."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop16."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D27', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E29', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB29', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH29', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN29', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E30', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB30', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH30', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN30', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E31', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB31', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH31', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN31', $poop24."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}
					//2 W, 1ero con 6 subw y el 2do con 5 subw
					if (count($g['retval']) == 3 && count($g['retval'][1]['subworks'])  == 30 && count($g['retval'][2]['subworks']) == 25) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->format('d-m-Y, g:i a'); ;
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->format('d-m-Y, g:i a'); ;
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->format('d-m-Y, g:i a'); ;
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->format('d-m-Y, g:i a'); ;
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$subwork15 = $g['retval'][1]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr15 = $dsr15->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][1]['subworks'][22]);
						$der15= $der15->format('d-m-Y, g:i a'); ;
						$poop15 = $g['retval'][1]['subworks'][23];
						$obs15 = $g['retval'][1]['subworks'][24];
						$subwork16 = $g['retval'][1]['subworks'][25];
						$dsr16 = new DateTime($g['retval'][1]['subworks'][26]);
						$dsr16 = $dsr16->format('d-m-Y, g:i a');
						$der16 = new DateTime($g['retval'][1]['subworks'][27]);
						$der16= $der16->format('d-m-Y, g:i a'); ;
						$poop16 = $g['retval'][1]['subworks'][28];
						$obs16 = $g['retval'][1]['subworks'][29];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->format('d-m-Y, g:i a'); ;
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->format('d-m-Y, g:i a'); ;
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->format('d-m-Y, g:i a'); ;
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->format('d-m-Y, g:i a'); ;
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];
						$subwork25 = $g['retval'][2]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][2]['subworks'][21]);
						$dsr25 = $dsr25->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][2]['subworks'][22]);
						$der25= $der25->format('d-m-Y, g:i a'); ;
						$poop25 = $g['retval'][2]['subworks'][23];
						$obs25 = $g['retval'][2]['subworks'][24];

						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E23', $subwork13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB23', $dsr13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH23', $der13);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN23', $poop13."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E24', $subwork14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB24', $dsr14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH24', $der14);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN24', $poop14."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E25', $subwork15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB25', $dsr15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH25', $der15);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN25', $poop15."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E26', $subwork16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB26', $dsr16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH26', $der16);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN26', $poop16."%");

						$objPHPExcel->getActiveSheet()->SetCellValue('D27', $work2);
						$objPHPExcel->getActiveSheet()->SetCellValue('E28', $subwork21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB28', $dsr21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH28', $der21);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN28', $poop21."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E29', $subwork22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB29', $dsr22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH29', $der22);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN29', $poop22."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E30', $subwork23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB30', $dsr23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH30', $der23);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN30', $poop23."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E31', $subwork24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB31', $dsr24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH31', $der24);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN31', $poop24."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E32', $subwork25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB32', $dsr25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH32', $der25);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN32', $poop25."%");

						header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$objWriter->save("php://output");
					}





				}


			}
		}









					/*for ($i=1; $i < count($g['retval']); $i++) {
						$work = $g['retval'][$i]['work'];
						echo $work1;
						for ($j=0; $j < count($g['retval'][$i]['subworks']); $j++) {
							$subwork = $g['retval'][$i]['subworks'][$j];
								echo $subwork;
								$dsr = $g['retval'][$i]['subworks'][$j];
								$der= $g['retval'][$i]['subworks'][$j];
								$poop = $g['retval'][$i]['subworks'][$j];
								$obs = $g['retval'][$i]['subworks'][$j];
								echo $dsr,$der,$poop,$obs;




						}
					}
					//var_dump($subwork12);

					/*foreach ($g as $v) {
						//$v = (array)$v;
					 	//$getwork[]=$v;
					 	var_dump($v['retval']['work']);
					}

					/*$filter  = array();
					$options = array('$sort' => array('works.work' => 1));
					$orderwork = $collwf->aggregate($options);
					$collobw = $db->orderbywork;
					$docworkobw = $collobw->insert($orderwork);*/
					//var_dump($orderwork['works']['work']);
					//$orderwork = iterator_to_array($orderwork,false);
					/*$cursor = $collwf->find();
					//$cursor = $collwf->findOne(["work" => "Cambio de caps 5" ]);
					$orderbywork = $cursor->sort(array('work' => 1));
					$works = $collwf->distinct("work");
					var_dump(count($works));
					foreach ($orderbywork as $v) {
						$v = (array)$v;
						 $getwork[]=$v;
					}
					$work = $getwork[1]['work'];
					$subwork = $getwork[1]['subwork'];
					var_dump($work,$subwork);//, $subwork;
					echo count($getwork);
					if ( count($works) == 1 ) {
						$array = array('work1' => $work,
							'subwork1.1' => $subwork
							);
						var_dump($docwork);
					}*/

					//return View::make('DataSend.report',array("docRepor" => $docRepor));
						/*$work_name = $w[0]['EQUIPMENT']['WORK']['WORK_NAME'];
						$subwork_name = $w['EQUIPMENT']['WORK']['SUBWORK']['SUBWORK_NAME'];
						$objPHPExcel->getActiveSheet()->SetCellValue('C'.(string)($row), $work_name);
						$objPHPExcel->getActiveSheet()->SetCellValue('C'.(string)($row), $subwork_name);



						/*$subwork_dsr = $arr['EQUIPMENT']['WORK']['SUBWORK']['DATE_START_REAL'];
						$subwork_der = $arr['EQUIPMENT']['WORK']['SUBWORK']['DATE_END_REAL'];
						$subwork_poop = $arr['EQUIPMENT']['WORK']['SUBWORK']['POOP'];
						$subwork_obs = $arr['EQUIPMENT']['WORK']['SUBWORK']['OBSERVATIONS'];
						$dsp = $arr['EQUIPMENT']['DATE_START_PROGRAMMED'];
						$dep = $arr['EQUIPMENT']['DATE_END_PROGRAMMED'];





						//$objPHPExcel->getActiveSheet()->getStyle('C'.(string)($row))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						$row ++;
						echo " exportando a excel ROW: ".$row;

					return View::make('DataSend.report',array("docRepor" => $docRepor));*/

						/*try {
						$objPHPExcel = $objReader->load("ReportOut.xlsx");
						} catch (Exception $e) {
							$objPHPExcel = $objReader->load("/var/www/senditlaravel42/ReportOut.xlsx");
						}
						$objPHPExcel->setActiveSheetIndex(0);
						$work = $objPHPExcel->getActiveSheet()->getCell('19')->getValue();

						/*$objPHPExcel->getActiveSheet()->SetCellValue('C'.(string)($row)), $names);
						$row ++;







						$docwork = $collwf->find(['EQUIPMENT.WORK.WORK_NAME' =>  $work_name]);
						$same_works = $db->$equal_works->insert($docwork);
						foreach ($variable as $key => $value) {

						}
						$update = $same_works->update(
							['Entry.AnswersJson.ADD_WORK_PAGE.WORK' => $work],
							[ '$set' => ['Entry.Id' => $IdForm]]

							);


						echo '<pre>'.var_dump($work_name,$subwork_name,$subwork_poop,$dsp,$dep).'</pre>';
						$work = new Work;
						$work->nombre = $work_name;
						$subwork = new Subwork;
						$subwork->works_id = $work_name;
						$subwork->nombre = $subwork_name;
						$subwork->fecha_inicio_real = $subwork_dsr;
						$subwork->fecha_termino_real = $subwork_der;
						$subwork->poop = $subwork_poop;
						$subwork->observaciones = $subwork_obs;

					if ($work->save() && $subwork->save()) {
						echo "insertado en work model and sub_work model";


			        }}
			        	$final_query = DB::table('subworks')->distinct()
									->select('subworks.nombre','subworks.works_id')
									->get();
			        	//var_dump($final_query);
			}
		}					//echo "insertado en work model and sub_work model";
	}						/*$first_query = DB::table('subworks')->distinct()
									->join('works','works.id','=','subworks.id')
									->select('subworks.nombre','subworks.fecha_inicio_real',
											'subworks.fecha_termino_real','subworks.poop',
											'subworks.observaciones');
							$query = Work::find(1)->subworks->id;
							var_dump($query);

							/*$final_query = DB::table('works')->distinct()
									->join('subworks','works.id','=','subworks.id')
									->select('works.nombre','subworks.fecha_inicio_real',
											'subworks.fecha_termino_real','subworks.poop',
											'subworks.observaciones')
									->get();
									var_dump($final_query->name->name->fecha_inicio_real->fecha_termino_real);

							$objPHPExcel = new PHPExcel();
							$objReader = PHPExcel_IOFactory::createReader('Excel2007');
							try {
								$objPHPExcel = $objReader->load("public/reports/reporteRudelEmpty.xlsx");
							} catch (Exception $e) {
								$objPHPExcel = $objReader->load("/var/www/senditlaravel42/public/reports/reporteRudelEmpty.xlsx");
							}
							$objPHPExcel->setActiveSheetIndex(0);

							/*foreach ($final_query as $q) {
								$names = $q->nombre;
								$dsr = $q->fecha_inicio_real;
								$der =	$q->fecha_termino_real;
								$poop =	$q->poop;
								$obs =$q->observaciones;
								var_dump($names, $dsr,$der,$poop,$obs);

								/*$objPHPExcel->getActiveSheet()->SetCellValue('C19', $names);
								$objPHPExcel->getActiveSheet()->SetCellValue('C19', $names);
								$objPHPExcel->getActiveSheet()->getStyle('C'.(string)($row))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
								$row ++;
							}*/





					//var_dump(iterator_to_array($docRepor,true));
					//$docReporArray = iterator_to_array($docRepor,false);
					/*for ($i=0; $i<count($docReporArray); $i++) {
						$work_name = $docReporArray[$i]['EQUIPMENT']['WORK']['WORK_NAME'];
						$work_name_next = $docReporArray[$i+1]['EQUIPMENT']['WORK']['WORK_NAME'];
						if ($work_name == $work_name_next ) {
							$work = new Work;
							$work->nombre = $work_name;
							$subwork = new Subwork;
							$subwork->nombre = $docReporArray[$i+1]['EQUIPMENT']['WORK']['SUBWORK']['SUBWORK_NAME'];
							if ($work->save() && $subwork->save()) {
								echo "trabajo insertado";
							}
						}
					}*//*foreach ($docReporArray as $key => $arr) {
						$work_name = $arr['EQUIPMENT']['WORK']['WORK_NAME'];
						$work_name_next = $arr['EQUIPMENT']['WORK']['WORK_NAME'];
						echo "{$key} => {$work_name} ";
						var_dump($docReporArray);
					}*/
					//echo "{$i} => {$work_name} ";
					//$key = array_search(40489, array_column($userdb, 'uid'));
					/*for ($i=0; $i<count($docReporArray); $i++) {
						//$keys = array_keys($docReporArray);
						$work_name = $docReporArray[$i]['EQUIPMENT']['WORK']['WORK_NAME'];
						$key = array_search($work_name, $docReporArray);




					}echo $key;*/
						//print $keys [1];
						//var_dump((array_keys($docReporArray)));



						/*$identificacion = $docRepor['EQUIPMENT']['IDENTIFICATION_EQUIPMENT']['IDENTIFICATION_NAME'];
						$localization = $docRepor['EQUIPMENT']['LOCALIZATION_EQUIPMENT']['LOCALIZATION_NAME'];

						$work = $docRepor['EQUIPMENT']['WORK']['WORK_NAME'];
						$sub_work = $docRepor['EQUIPMENT']['WORK']['SUBWORK']['SUBWORK_NAME'];
						$date_start_real = $docRepor['EQUIPMENT']['WORK']['SUBWORK']['DATE_START_REAL'];
						$date_end_real = $docRepor['EQUIPMENT']['WORK']['SUBWORK']['DATE_END_REAL'];
						$poop = $docRepor['EQUIPMENT']['WORK']['SUBWORK']['POOP'];
						$obs = $docRepor['EQUIPMENT']['WORK']['SUBWORK']['OBSERVATIONS'];

					}


					//return View::make('DataSend.report',array("docRepor" => $docRepor));

				}

		//return View::make('DataSend.report',array("docRepor" => $docRepor));



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

		$aRequest = json_decode(file_get_contents('php://input'),true);
		$fichero=fopen('test.log','w');
	 		if($fichero == false) {
   			die("No se ha podido crear el archivo.");
		}
		fwrite($fichero,json_encode($aRequest));
		fclose($fichero);



		$m = new MongoClient();//obsoleta desde mongo 1.0.0
		$db = $m->SenditForm;
		$collRepor = $db->Repor;
		if ($collRepor->count() == 0 ){
				$array = array(
					"ProviderId" => $aRequest['ProviderId'],
					"IntegrationKey" => $aRequest['IntegrationKey'],
					"Entry" => array(
						 "Id" => $aRequest['Entry']['Id'],
						 "FormCode" => $aRequest['Entry']['FormCode'],
						 "FormVersion" => $aRequest['Entry']['FormVersion'],
						 "UserFirstName" => $aRequest['Entry']['UserFirstName'],
						 "UserLastName" => $aRequest['Entry']['UserLastName'],
						 "UserEmail" => $aRequest['Entry']['UserEmail'],
						 "Latitude" => $aRequest['Entry']['Latitude'],
						 "Longitude" => $aRequest['Entry']['Longitude'],
						 "StartTime" => $aRequest['Entry']['StartTime'],
						 "ReceivedTime" => $aRequest['Entry']['ReceivedTime'],
						 "CompleteTime" => $aRequest['Entry']['CompleteTime']
						),
					"EQUIPMENT" => array(
						"EQUIPMENT_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['EQUIPMENT'],
						"IDENTIFICATION_EQUIPMENT" => array(
							"IDENTIFICATION_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['IDENTIFICATION_EQUIPMENT']
							),
						"LOCALIZATION_EQUIPMENT" => array(
							"LOCALIZATION_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['LOCALIZATION_EQUIPMENT']
							),
						"BLOCK_SYSTEM" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['BLOCK_SYSTEM'],
						"DATE_START_PROGRAMMED" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_PROGRAMMED'],
						"DATE_END_PROGRAMMED" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_PROGRAMMED'],
						"WORK" =>  array(
							"WORK_NUEVO" => "SI",
							"WORK_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['WORK'],
							"SUBWORK" => array(
								"SUBWORK_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['SUBWORK'],
								"DATE_START_REAL" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_REAL'],
								"DATE_END_REAL" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_REAL'],
								"POOP" => "60",
								"OBSERVATIONS" =>  $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['OBSERVATIONS']
								),
							"TURNS_PAGE" => array(
								"S_TURN_DAY" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['S_TURN_DAY'],
						  		"S_TURN_NIGHT" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['S_TURN_NIGHT'],
						  		"I_P_TURN_DAY" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['I_P_TURN_DAY'],
						  		"I_P_TURN_NIGHT" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['I_P_TURN_NIGHT']
							),
							"PHOTOS" => array(
								"PHOTO1" => $aRequest['Entry']['AnswersJson']['PHOTOS']['PHOTO1'],
								"DESCRIPTION_PHOTO1" => $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_PHOTO1'],
								"PHOTO2" => $aRequest['Entry']['AnswersJson']['PHOTOS']['PHOTO2'],
								"DESCRIPTION_PHOTO2" => $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_PHOTO2'],
								"PHOTO3" => $aRequest['Entry']['AnswersJson']['PHOTOS']['PHOTO3'],
								"DESCRIPTION_PHOTO3" => $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_PHOTO3'],
								"VIDEO" => $aRequest['Entry']['AnswersJson']['PHOTOS']['VIDEO'],
								"DESCRIPTION_VIDEO"=> $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_VIDEO'],
								"NEXT_PAGE_FORM_P" => $aRequest['Entry']['AnswersJson']['PHOTOS']['NEXT_PAGE_FORM_P']
							)
						)
					)
				);
				$docRepor = $collRepor->insert($array);
				echo "Insertado en Repor work nuevo : si";
			/*if ($result) {
				$array = array(
					"ProviderId" => $aRequest['ProviderId'],
					"IntegrationKey" => $aRequest['IntegrationKey'],
					"Entry" => array(
						 "Id" => $aRequest['Entry']['Id'],
						 "FormCode" => $aRequest['Entry']['FormCode'],
						 "FormVersion" => $aRequest['Entry']['FormVersion'],
						 "UserFirstName" => $aRequest['Entry']['UserFirstName'],
						 "UserLastName" => $aRequest['Entry']['UserLastName'],
						 "UserEmail" => $aRequest['Entry']['UserEmail'],
						 "Latitude" => $aRequest['Entry']['Latitude'],
						 "Longitude" => $aRequest['Entry']['Longitude'],
						 "StartTime" => $aRequest['Entry']['StartTime'],
						 "ReceivedTime" => $aRequest['Entry']['ReceivedTime'],
						 "CompleteTime" => $aRequest['Entry']['CompleteTime']
						),
					"EQUIPMENT" => array(
						"EQUIPMENT_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['EQUIPMENT'],
						"IDENTIFICATION_EQUIPMENT" => array(
							"IDENTIFICATION_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['IDENTIFICATION_EQUIPMENT']
							),
						"LOCALIZATION_EQUIPMENT" => array(
							"LOCALIZATION_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['LOCALIZATION_EQUIPMENT']
							),
						"BLOCK_SYSTEM" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['BLOCK_SYSTEM'],
						"DATE_START_PROGRAMMED" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_PROGRAMMED'],
						"DATE_END_PROGRAMMED" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_PROGRAMMED'],
						"WORK" =>  array(
							"WORK_NUEVO" => "NO",
							"WORK_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['WORK'],
							"SUBWORK" => array(
								"SUBWORK_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['SUBWORK'],
								"DATE_START_REAL" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_REAL'],
								"DATE_END_REAL" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_REAL'],
								"POOP" => "60",
								"OBSERVATIONS" =>  $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['OBSERVATIONS']
								),
							"TURNS_PAGE" => array(
								"S_TURN_DAY" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['S_TURN_DAY'],
						  		"S_TURN_NIGHT" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['S_TURN_NIGHT'],
						  		"I_P_TURN_DAY" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['I_P_TURN_DAY'],
						  		"I_P_TURN_NIGHT" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['I_P_TURN_NIGHT']
							),
							"PHOTOS" => array(
								"PHOTO1" => $aRequest['Entry']['AnswersJson']['PHOTOS']['PHOTO1'],
								"DESCRIPTION_PHOTO1" => $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_PHOTO1'],
								"PHOTO2" => $aRequest['Entry']['AnswersJson']['PHOTOS']['PHOTO2'],
								"DESCRIPTION_PHOTO2" => $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_PHOTO2'],
								"PHOTO3" => $aRequest['Entry']['AnswersJson']['PHOTOS']['PHOTO3'],
								"DESCRIPTION_PHOTO3" => $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_PHOTO3'],
								"VIDEO" => $aRequest['Entry']['AnswersJson']['PHOTOS']['VIDEO'],
								"DESCRIPTION_VIDEO"=> $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_VIDEO'],
								"NEXT_PAGE_FORM_P" => $aRequest['Entry']['AnswersJson']['PHOTOS']['NEXT_PAGE_FORM_P']
							)
						)
						)
				);
				$docRepor = $collRepor->insert($array);
				echo "Insertado en Repor work nuevo : no";
			}else{

			}*/
		}else{
			echo "no vacio";
			$array = array(
					"ProviderId" => $aRequest['ProviderId'],
					"IntegrationKey" => $aRequest['IntegrationKey'],
					"Entry" => array(
						 "Id" => $aRequest['Entry']['Id'],
						 "FormCode" => $aRequest['Entry']['FormCode'],
						 "FormVersion" => $aRequest['Entry']['FormVersion'],
						 "UserFirstName" => $aRequest['Entry']['UserFirstName'],
						 "UserLastName" => $aRequest['Entry']['UserLastName'],
						 "UserEmail" => $aRequest['Entry']['UserEmail'],
						 "Latitude" => $aRequest['Entry']['Latitude'],
						 "Longitude" => $aRequest['Entry']['Longitude'],
						 "StartTime" => $aRequest['Entry']['StartTime'],
						 "ReceivedTime" => $aRequest['Entry']['ReceivedTime'],
						 "CompleteTime" => $aRequest['Entry']['CompleteTime']
						),
					"EQUIPMENT" => array(
						"EQUIPMENT_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['EQUIPMENT'],
						"IDENTIFICATION_EQUIPMENT" => array(
							"IDENTIFICATION_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['IDENTIFICATION_EQUIPMENT']
							),
						"LOCALIZATION_EQUIPMENT" => array(
							"LOCALIZATION_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['LOCALIZATION_EQUIPMENT']
							),
						"BLOCK_SYSTEM" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['BLOCK_SYSTEM'],
						"DATE_START_PROGRAMMED" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_PROGRAMMED'],
						"DATE_END_PROGRAMMED" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_PROGRAMMED'],
						"WORK" =>  array(
							"WORK_NUEVO" => "NO",
							"WORK_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['WORK'],
							"SUBWORK" => array(
								"SUBWORK_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['SUBWORK'],
								"DATE_START_REAL" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_REAL'],
								"DATE_END_REAL" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_REAL'],
								"POOP" => "60",
								"OBSERVATIONS" =>  $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['OBSERVATIONS']
								),
							"TURNS_PAGE" => array(
								"S_TURN_DAY" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['S_TURN_DAY'],
						  		"S_TURN_NIGHT" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['S_TURN_NIGHT'],
						  		"I_P_TURN_DAY" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['I_P_TURN_DAY'],
						  		"I_P_TURN_NIGHT" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['I_P_TURN_NIGHT']
							),
							"PHOTOS" => array(
								"PHOTO1" => $aRequest['Entry']['AnswersJson']['PHOTOS']['PHOTO1'],
								"DESCRIPTION_PHOTO1" => $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_PHOTO1'],
								"PHOTO2" => $aRequest['Entry']['AnswersJson']['PHOTOS']['PHOTO2'],
								"DESCRIPTION_PHOTO2" => $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_PHOTO2'],
								"PHOTO3" => $aRequest['Entry']['AnswersJson']['PHOTOS']['PHOTO3'],
								"DESCRIPTION_PHOTO3" => $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_PHOTO3'],
								"VIDEO" => $aRequest['Entry']['AnswersJson']['PHOTOS']['VIDEO'],
								"DESCRIPTION_VIDEO"=> $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_VIDEO'],
								"NEXT_PAGE_FORM_P" => $aRequest['Entry']['AnswersJson']['PHOTOS']['NEXT_PAGE_FORM_P']
							)
						)
						)
				);
			$result = $collRepor->findOne([
				'EQUIPMENT.WORK.WORK_NAME' => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['WORK']
				]);
			//$result = (array)$result;

				//var_dump($value['EQUIPMENT']['WORK']['WORK_NAME']);
			if ($result) {

			$docRepor = $collRepor->insert($array);
			echo "Insertado en Repor work nuevo : no";
			}else{
				$array = array(
					"ProviderId" => $aRequest['ProviderId'],
					"IntegrationKey" => $aRequest['IntegrationKey'],
					"Entry" => array(
						 "Id" => $aRequest['Entry']['Id'],
						 "FormCode" => $aRequest['Entry']['FormCode'],
						 "FormVersion" => $aRequest['Entry']['FormVersion'],
						 "UserFirstName" => $aRequest['Entry']['UserFirstName'],
						 "UserLastName" => $aRequest['Entry']['UserLastName'],
						 "UserEmail" => $aRequest['Entry']['UserEmail'],
						 "Latitude" => $aRequest['Entry']['Latitude'],
						 "Longitude" => $aRequest['Entry']['Longitude'],
						 "StartTime" => $aRequest['Entry']['StartTime'],
						 "ReceivedTime" => $aRequest['Entry']['ReceivedTime'],
						 "CompleteTime" => $aRequest['Entry']['CompleteTime']
						),
					"EQUIPMENT" => array(
						"EQUIPMENT_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['EQUIPMENT'],
						"IDENTIFICATION_EQUIPMENT" => array(
							"IDENTIFICATION_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['IDENTIFICATION_EQUIPMENT']
							),
						"LOCALIZATION_EQUIPMENT" => array(
							"LOCALIZATION_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['LOCALIZATION_EQUIPMENT']
							),
						"BLOCK_SYSTEM" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['BLOCK_SYSTEM'],
						"DATE_START_PROGRAMMED" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_PROGRAMMED'],
						"DATE_END_PROGRAMMED" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_PROGRAMMED'],
						"WORK" =>  array(
							"WORK_NUEVO" => "SI",
							"WORK_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['WORK'],
							"SUBWORK" => array(
								"SUBWORK_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['SUBWORK'],
								"DATE_START_REAL" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_REAL'],
								"DATE_END_REAL" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_REAL'],
								"POOP" => "60",
								"OBSERVATIONS" =>  $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['OBSERVATIONS']
								),
							"TURNS_PAGE" => array(
								"S_TURN_DAY" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['S_TURN_DAY'],
						  		"S_TURN_NIGHT" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['S_TURN_NIGHT'],
						  		"I_P_TURN_DAY" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['I_P_TURN_DAY'],
						  		"I_P_TURN_NIGHT" => $aRequest['Entry']['AnswersJson']['TURNS_PAGE']['I_P_TURN_NIGHT']
							),
							"PHOTOS" => array(
								"PHOTO1" => $aRequest['Entry']['AnswersJson']['PHOTOS']['PHOTO1'],
								"DESCRIPTION_PHOTO1" => $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_PHOTO1'],
								"PHOTO2" => $aRequest['Entry']['AnswersJson']['PHOTOS']['PHOTO2'],
								"DESCRIPTION_PHOTO2" => $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_PHOTO2'],
								"PHOTO3" => $aRequest['Entry']['AnswersJson']['PHOTOS']['PHOTO3'],
								"DESCRIPTION_PHOTO3" => $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_PHOTO3'],
								"VIDEO" => $aRequest['Entry']['AnswersJson']['PHOTOS']['VIDEO'],
								"DESCRIPTION_VIDEO"=> $aRequest['Entry']['AnswersJson']['PHOTOS']['DESCRIPTION_VIDEO'],
								"NEXT_PAGE_FORM_P" => $aRequest['Entry']['AnswersJson']['PHOTOS']['NEXT_PAGE_FORM_P']
							)
						)
					)
				);
				$docRepor = $collRepor->insert($array);
				echo "Insertado en Repor work nuevo : SI";
			}

			//var_dump($result);


		}

	}//class store

		/*
		$Work = $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['WORK'];
		$Sub_W = $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['SUBWORK'];
		$Work = $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['WORK'];
		$Sub_W = $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['SUBWORK'];
		/*if ($collection->count() > 0) {
			$docSendit = $collection->find(['Entry.AnswersJson.Trabajos_planificados2.Trabajos' => $Work]);
			if ($docSendit->count() > 0) {
				# update
				//$newSubW = $aRequest($set => ['Entry.AnswersJson.Trabajos_planificados2.Trabajos'] => $Sub_W );
				//$updateResult = $collection->update(['Entry.AnswersJson.Trabajos_planificados2.Trabajos' => $Work], $newSubW);
				/*$updateResult = $collection->update(
			    ['Entry.AnswersJson.Trabajos_planificados2.Trabajos' => $Work],
			    [ '$set' => ['Entry.AnswersJson.Trabajos_planificados2.Sub_trabajo' => $Sub_W]]
			);

				var_dump($collection->find(['Entry.AnswersJson.Trabajos_planificados2.Trabajos' => $Work]));
				echo "Trabajo updated";

			}else{
				$doc = $collection->insert($aRequest);
		 		echo "Trabajo nuevo insertado";
			}
		}else{
			$doc = $collection->insert($aRequest);
		 	echo "Coleccion vacia nuevo trabajo insertado";
		}
		$docWork = $collWorks->findOne(['Entry.AnswersJson.ADD_WORK_PAGE.WORK' => $Work]);
		//var_dump($docWork);
		$work = $docWork['Entry']['AnswersJson']['ADD_WORK_PAGE']['WORK'];
		//echo $work;
		if ($docWork){
			# code...
			//$docSendit = $collection->find(['Entry.AnswersJson.Trabajos_planificados2.Trabajos' => $Work]);
			$IdForm = $docWork['Entry']['Id'];//get id de Works collection
			$collSubWorks = $db->SubWorks;//create collection

			$docSubWs = $collSubWorks->insert($aRequest);
			//$subws = $collSubWorks->find();
			$updateResult = $collSubWorks->update(
				['Entry.AnswersJson.ADD_WORK_PAGE.WORK' => $work],
				[ '$set' => ['Entry.Id' => $Id(Form]],
								['multiple' => true]
							);
							/*foreach ($subws as $subw) {
								$updateResult = $subw->update(
							    ['Entry.AnswersJson.Trabajos_planificados2.Trabajos' => $work],
							    [ '$set' => ['Entry.Id' => $IdForm]]
							);)
			$work = $subw['Entry']['AnswersJson']['Trabajos_planificados2']['Trabajos'];
			echo $work;
			//echo $subW->Entry->AnswersJson->Trabajos_planificados2->Trabajos;

			}
			/*for($i=0;$i<count($subws);$i++){
				$id_fruta=$subws[$i]->Entry->AnswersJson->Trabajos_planificados2->Trabajos;

			    echo $id_fruta;
			}

			echo "doc insertado en SubWs collection";
		}else{

		$docWork = $collWorks->insert($aRequest);
		//$indexName = $collection->createIndex(['borough' => 1, 'cuisine' => 1]);
		echo "doc insertado en Works collection";

		}
		//$doc = $collection->insert($aRequest);
		//echo "doc insertado";
		 //$email = $aRequest['Entry']['UserEmail'];
		 //$email = $aRequest['Entry']['UserEmail'];
		//echo $email;
		/*$providerId = $aRequest['ProviderId'];//id del proveedor del json entrante
		$docSendit = $collection->findOne(['ProviderId' => $providerId]);
		echo "mostrando email de la bdmongo: ".$docSendit['Entry']['UserEmail'];
		echo $docSendit['Entry']['UserFirstName'];
		echo $docSendit['Entry']['UserLastName'];
		echo $docSendit['Entry']['Latitude'];
		$email = $collection->findOne(['Entry.UserEmail' => $email]);*/

		//var_dump($providerId);
		//var_dump($email);
		//printf("Inserted %d documents",$insert->getInsertedCount());
		 //echo "mostrando id deldocumento insertado";
		//var_dump($doc->getInsertedId());

	/*foreach ($db->listCollections() as $collec) {
			# code...
			echo "mostrando colecciones";
			var_dump($collec);
		}*/



	/**
	 * Display the specified resource.
	 *
	 * @param  int  $id
	 * @return Response
	 */
	public function show($id)
	{
		//
		return "hola soy show" .$id;
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
