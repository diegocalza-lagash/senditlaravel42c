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

	/*public function __construct(){
		$this->beforeFilter('auth.user');
	}*/
	public function getEquipments(){
		echo "hola";

	}

	public function getIndex()
	{

		$this->layout->content = View::make('DataSend.index');
		//return View::make('DataSend.report',array("docRepor" => $docRepor));
	}

	public function report(){
		//echo "hola";
		if (isset($_GET["equi"]) /*&& isset($_POST['loc']) && isset($_POST['iden']) && isset($_POST['dsp']) && isset($_POST['dep'])*/){
			$equi = htmlspecialchars(Input::get("equi"));
			$loc = htmlspecialchars(Input::get("loc"));
			$iden = htmlspecialchars(Input::get("iden"));
			$dsp = htmlspecialchars(Input::get("dsp"));
			//$dsp = $_GET['dsp'];
			//echo $dsp;
			$dep = htmlspecialchars(Input::get("dep"));
			/*$equi = $_POST["equi"];
			$loc = $_POST["loc"];echo $loc;
			$iden = $_POST['iden'];
			$dsp = $_POST['dsp'];
			$dep = $_POST['dep'];*/
				/*$loc = htmlspecialchars(Input::get("loc"));
				$iden = htmlspecialchars(Input::get("iden"));
				$dsp = htmlspecialchars(Input::get("dsp"));
				$dep = htmlspecialchars(Input::get("dep"));*/
			$dsp = new DateTime($dsp);
			//$dsp->setTimezone(new DateTimeZone('America/Santiago'));
			$dsp = $dsp->format('d/m/Y');
			$dep = new DateTime($dep);
			$dep = $dep->format('d/m/Y');
			//echo $dsp." ". $dep;
			$m = new MongoClient();//obsoleta desde mongo 1.0.0
			$db = $m->SenditForm;
			$collRepor = $db->Repor;
			//echo $equi." ".$loc." ".$iden." ".$dsp." ". $dep;
			$docRepor = $collRepor->find([
				'EQUIPMENT.EQUIPMENT_NAME' => $equi,
				'EQUIPMENT.LOCALIZATION_EQUIPMENT.LOCALIZATION_NAME' => $loc,
				'EQUIPMENT.IDENTIFICATION_EQUIPMENT.IDENTIFICATION_NAME' => $iden,
				'EQUIPMENT.DATE_START_PROGRAMMED' => $dsp,
				'EQUIPMENT.DATE_END_PROGRAMMED' => $dep
				]);

				//return View::make('DataSend.report',array("docRepor" => $docRepor));
				if (!$docRepor -> count()) {
					//Session::flash('mensaje_error', 'No Existen Trabajos')
					return Redirect::to('/dataform')
                    ->with('mensaje_error', 'No Existen Trabajos');
				}else{
					$m = new MongoClient();
					$db = $m->SenditForm;
					$collwf = $db->works_filter;

					if ($collwf->count()>0) {
						$collwf->drop();
					}
					foreach ($docRepor as  $v) {

						$docwork = $collwf->insert(array(
							"work" => $v["EQUIPMENT"]["WORK"]["WORK_NAME"],
							"subwork" => array(
								"subw_name" => $v["EQUIPMENT"]["WORK"]["SUBWORK"]["SUBWORK_NAME"],
								"work_nuevo" => $v['EQUIPMENT']['WORK']['WORK_NUEVO'],
								"dsr" => $v['EQUIPMENT']['WORK']['SUBWORK']['DATE_START_REAL'],
								"der" => $v['EQUIPMENT']['WORK']['SUBWORK']['DATE_END_REAL'],
								"poop" => $v['EQUIPMENT']['WORK']['SUBWORK']['POOP'],
								"obs" => $v['EQUIPMENT']['WORK']['SUBWORK']['OBSERVATIONS']
								)
							));

					}
					//imprimo los datos fijos del reporte
					foreach ($docRepor as $v) {
						$std = $v['EQUIPMENT']['WORK']['TURNS_PAGE']['S_TURN_DAY'];
						$stn = $v['EQUIPMENT']['WORK']['TURNS_PAGE']['S_TURN_NIGHT'];
						$iptd = $v['EQUIPMENT']['WORK']['TURNS_PAGE']['I_P_TURN_DAY'];
						$iptn = $v['EQUIPMENT']['WORK']['TURNS_PAGE']['I_P_TURN_NIGHT'];
						$block = $v['EQUIPMENT']['BLOCK_SYSTEM'];
						$fip = $v['EQUIPMENT']['DATE_START_PROGRAMMED'];
						$ftp = $v['EQUIPMENT']['DATE_END_PROGRAMMED'];
						$hp = $v['EQUIPMENT']['HOUR_PROG'];
						}//endforeach

						$objPHPExcel = new PHPExcel();
						$objReader = PHPExcel_IOFactory::createReader('Excel2007');
						try {
							$objPHPExcel = $objReader->load("public/reporteRudel.xlsx");
						} catch (Exception $e) {
							$objPHPExcel = $objReader->load("/var/www/senditlaravel42/public/reporteRudel.xlsx");
						}

						$objWorksheet= $objPHPExcel->setActiveSheetIndex(0);
						$objPHPExcel->getActiveSheet()->SetCellValue('H9', $loc);
						$objPHPExcel->getActiveSheet()->SetCellValue('H11', $block);
						$objPHPExcel->getActiveSheet()->SetCellValue('I14', $fip);
						$objPHPExcel->getActiveSheet()->SetCellValue('I15', $ftp);
						$objPHPExcel->getActiveSheet()->SetCellValue('I16', $hp);
						$objPHPExcel->getActiveSheet()->SetCellValue('AD9', $std);
						$objPHPExcel->getActiveSheet()->SetCellValue('AD10', $stn);
						$objPHPExcel->getActiveSheet()->SetCellValue('AD11', $iptd);
						$objPHPExcel->getActiveSheet()->SetCellValue('AD12', $iptn);
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						try {
							$objWriter->save("public/reporteRudel2.xlsx");
						} catch (Exception $e) {
							$objWriter->save("/var/www/senditlaravel42/public/reporteRudel2.xlsx");
						}

					$keys = array("work" => 1);
					$initial = array("subworks" => array());//,obj.std,obj.stn,obj.iptd,obj.iptn,obj.fip,obj.ftp,obj.hp
					$reduce = "function(obj, prev){
						prev.subworks.push(obj.subwork.subw_name,obj.subwork.dsr,obj.subwork.der,obj.subwork.poop,obj.subwork.obs)
					}";
					$g = $collwf->group($keys,$initial,$reduce);
					//$collwf->drop();
					$collwf = $db->works_filter;
					$docwork = $collwf->insert($g);
					$objPHPExcel = new PHPExcel();
					$objReader = PHPExcel_IOFactory::createReader('Excel2007');
					try {
						$objPHPExcel = $objReader->load("public/reporteRudel2.xlsx");
					} catch (Exception $e) {
						$objPHPExcel = $objReader->load("/var/www/senditlaravel42/public/reporteRudel2.xlsx");
					}
					$objWorksheet= $objPHPExcel->setActiveSheetIndex(0);
					//echo count($g);
					//echo json_encode($g['retval']);

					//1w y 1
					if ($g['retval'][0]['work'] != null && count($g['retval']) == 1 && count($g['retval'][0]['subworks']) == 5) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						//var_dump(count($g['retval'][1]['subworks']));
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						//echo $dsr11."".$der11;

						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						//$collwf->drop();//para que no agrege null en los works
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//1w  y 2
					if (count($g['retval']) == 1 && count($g['retval'][0]['subworks']) == 10) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						//$dsr12 = $dsr12;
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12 = $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						//$der12 = $der12
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$objPHPExcel->getActiveSheet()->SetCellValue('D20', $work1);
						$objPHPExcel->getActiveSheet()->SetCellValue('E21', $subwork11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB21', $dsr11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH21', $der11);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN21', $poop11."%");
						$objPHPExcel->getActiveSheet()->SetCellValue('E22', $subwork12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AB22', $dsr12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AH22', $der12);
						$objPHPExcel->getActiveSheet()->SetCellValue('AN22', $poop12."%");

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");

						return View::make('DataSend.report', array("docRepor" => $docRepor));

					}
					//1w  y 3
					if (count($g['retval']) == 1 && count($g['retval'][0]['subworks']) == 15) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12 = $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13 = $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));

					}
					// 1w  y 4
					if (count($g['retval']) == 1 && count($g['retval'][0]['subworks']) == 20) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12 = $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13 = $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14 = $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];
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

						//header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
						//header('Content-Disposition: attachment; filename="ReportOut.xlsx"');
						//header("Cache-Control: max-age=0");
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						//$objWriter->save("php://output");
						return View::make('DataSend.report', array("docRepor" => $docRepor));

					}
					//1w  y 5
					if (count($g['retval']) == 1 && count($g['retval'][0]['subworks']) == 25) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12 = $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13 = $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14 = $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];
						$subwork15 = $g['retval'][0]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][0]['subworks'][21]);
						$dsr15 = $dsr15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][0]['subworks'][22]);
						$der15 = $der15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][0]['subworks'][23];
						$obs15 = $g['retval'][0]['subworks'][24];
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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));

					}
					//1w  y 6
					if (count($g['retval']) == 1 && count($g['retval'][0]['subworks']) == 30) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12 = $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13 = $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14 = $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];
						$subwork15 = $g['retval'][0]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][0]['subworks'][21]);
						$dsr15 = $dsr15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][0]['subworks'][22]);
						$der15 = $der15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][0]['subworks'][23];
						$obs15 = $g['retval'][0]['subworks'][24];
						$subwork16 = $g['retval'][0]['subworks'][25];
						$dsr16 = new DateTime($g['retval'][0]['subworks'][26]);
						$dsr16 = $dsr16->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der16 = new DateTime($g['retval'][0]['subworks'][27]);
						$der16 = $der16->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop16 = $g['retval'][0]['subworks'][28];
						$obs16 = $g['retval'][0]['subworks'][29];
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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));

					}
					//ASCENDENTE
					//2w 1 y 1
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks']) == 5 && count($g['retval'][1]['subworks']) == 5) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2w 1 y 2 **** desc ya esta
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 5 && count($g['retval'][1]['subworks']) == 10) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');;
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2w 1 y 3
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks']) == 5 && count($g['retval'][1]['subworks']) == 15) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2w 1 y 4
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks']) == 5 && count($g['retval'][1]['subworks']) == 20) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24 = $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];

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
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2w 1 y 5
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks']) == 5 && count($g['retval'][1]['subworks']) == 25) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24 = $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];
						$subwork25 = $g['retval'][1]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr25 = $dsr25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][1]['subworks'][22]);
						$der25 = $der25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop25 = $g['retval'][1]['subworks'][23];
						$obs25 = $g['retval'][1]['subworks'][24];
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
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2w 1 y 6
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks']) == 5 && count($g['retval'][1]['subworks']) == 30) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24 = $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];
						$subwork25 = $g['retval'][1]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr25 = $dsr25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][1]['subworks'][22]);
						$der25 = $der25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop25 = $g['retval'][1]['subworks'][23];
						$obs25 = $g['retval'][1]['subworks'][24];
						$subwork26 = $g['retval'][1]['subworks'][25];
						$dsr26 = new DateTime($g['retval'][1]['subworks'][26]);
						$dsr26 = $dsr26->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der26 = new DateTime($g['retval'][1]['subworks'][27]);
						$der26 = $der26->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop26 = $g['retval'][1]['subworks'][28];
						$obs26 = $g['retval'][1]['subworks'][29];
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
						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2W 1ero 2SUB
					//2 W, 2 y 2
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 10 && count($g['retval'][1]['subworks']) == 10) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 2 y 3
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 10 && count($g['retval'][1]['subworks']) == 15) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 2 y 4*
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 10 && count($g['retval'][1]['subworks']) == 20) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24= $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 2 y 5
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 10 && count($g['retval'][1]['subworks']) == 25) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24= $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];
						$subwork25 = $g['retval'][1]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr25 = $dsr25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][1]['subworks'][22]);
						$der25= $der25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop25 = $g['retval'][1]['subworks'][23];
						$obs25 = $g['retval'][1]['subworks'][24];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 2 y 6
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 10 && count($g['retval'][1]['subworks']) == 30) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24= $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];
						$subwork25 = $g['retval'][1]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr25 = $dsr25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][1]['subworks'][22]);
						$der25= $der25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop25 = $g['retval'][1]['subworks'][23];
						$obs25 = $g['retval'][1]['subworks'][24];
						$subwork26 = $g['retval'][1]['subworks'][25];
						$dsr26 = new DateTime($g['retval'][1]['subworks'][26]);
						$dsr26 = $dsr26->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der26 = new DateTime($g['retval'][1]['subworks'][27]);
						$der26= $der26->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop26 = $g['retval'][1]['subworks'][28];
						$obs26 = $g['retval'][1]['subworks'][29];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2W 1ero 3SUB
					//2 W, 3 y 1
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 15 && count($g['retval'][1]['subworks']) == 5) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 3 y 2
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 15 && count($g['retval'][1]['subworks']) == 10) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 3 y 3
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 15 && count($g['retval'][1]['subworks']) == 15) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 3 y 4
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 15 && count($g['retval'][1]['subworks']) == 20) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24= $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 3 y 5
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 15 && count($g['retval'][1]['subworks']) == 25) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24= $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];
						$subwork25 = $g['retval'][1]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr25 = $dsr25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][1]['subworks'][22]);
						$der25= $der25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop25 = $g['retval'][1]['subworks'][23];
						$obs25 = $g['retval'][1]['subworks'][24];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 3 y 6
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 15 && count($g['retval'][1]['subworks']) == 30) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24= $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];
						$subwork25 = $g['retval'][1]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr25 = $dsr25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][1]['subworks'][22]);
						$der25= $der25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop25 = $g['retval'][1]['subworks'][23];
						$obs25 = $g['retval'][1]['subworks'][24];
						$subwork26 = $g['retval'][1]['subworks'][25];
						$dsr26 = new DateTime($g['retval'][1]['subworks'][26]);
						$dsr26 = $dsr26->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der26 = new DateTime($g['retval'][1]['subworks'][27]);
						$der26= $der26->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop26 = $g['retval'][1]['subworks'][28];
						$obs26 = $g['retval'][1]['subworks'][29];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2W 1ero 4Sub
					//2 W, 4 y 4
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 20 && count($g['retval'][1]['subworks']) == 20) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24= $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];

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


						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 4 y 5
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 20 && count($g['retval'][1]['subworks']) == 25) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->fsetTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24= $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];
						$subwork25 = $g['retval'][1]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr25 = $dsr25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][1]['subworks'][22]);
						$der25= $der25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop25 = $g['retval'][1]['subworks'][23];
						$obs25 = $g['retval'][1]['subworks'][24];

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


						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 4 y 6
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 20 && count($g['retval'][1]['subworks']) == 30) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24= $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];
						$subwork25 = $g['retval'][1]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr25 = $dsr25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][1]['subworks'][22]);
						$der25= $der25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop25 = $g['retval'][1]['subworks'][23];
						$obs25 = $g['retval'][1]['subworks'][24];
						$subwork26 = $g['retval'][1]['subworks'][25];
						$dsr26 = new DateTime($g['retval'][1]['subworks'][26]);
						$dsr26 = $dsr26->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der26 = new DateTime($g['retval'][1]['subworks'][27]);
						$der26= $der26->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop26 = $g['retval'][1]['subworks'][28];
						$obs26 = $g['retval'][1]['subworks'][29];


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


						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2W 1ero 5Sub
					//2 W, 5 y 5
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 25 && count($g['retval'][1]['subworks']) == 25) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];
						$subwork15 = $g['retval'][0]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][0]['subworks'][21]);
						$dsr15 = $dsr15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][0]['subworks'][22]);
						$der15= $der15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][0]['subworks'][23];
						$obs15 = $g['retval'][0]['subworks'][24];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24= $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];
						$subwork25 = $g['retval'][1]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr25 = $dsr25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][1]['subworks'][22]);
						$der25= $der25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop25 = $g['retval'][1]['subworks'][23];
						$obs25 = $g['retval'][1]['subworks'][24];


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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 5 y 6
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 25 && count($g['retval'][1]['subworks']) == 30) {
						$work1 = $g['retval'][1]['work'];
						$subwork11 = $g['retval'][1]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][1]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][1]['subworks'][3];
						$obs11 = $g['retval'][1]['subworks'][4];
						$subwork12 = $g['retval'][1]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][1]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][1]['subworks'][8];
						$obs12 = $g['retval'][1]['subworks'][9];
						$subwork13 = $g['retval'][1]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][1]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][1]['subworks'][13];
						$obs13 = $g['retval'][1]['subworks'][14];
						$subwork14 = $g['retval'][1]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][1]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][1]['subworks'][18];
						$obs14 = $g['retval'][1]['subworks'][19];
						$subwork15 = $g['retval'][1]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr15 = $dsr15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][1]['subworks'][22]);
						$der15= $der15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][1]['subworks'][23];
						$obs15 = $g['retval'][1]['subworks'][24];

						$work2 = $g['retval'][2]['work'];
						$subwork21 = $g['retval'][2]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][2]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][2]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][2]['subworks'][3];
						$obs21 = $g['retval'][2]['subworks'][4];
						$subwork22 = $g['retval'][2]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][2]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][2]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][2]['subworks'][8];
						$obs22 = $g['retval'][2]['subworks'][9];
						$subwork23 = $g['retval'][2]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][2]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][2]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][2]['subworks'][13];
						$obs23 = $g['retval'][2]['subworks'][14];
						$subwork24 = $g['retval'][2]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][2]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][2]['subworks'][17]);
						$der24= $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][2]['subworks'][18];
						$obs24 = $g['retval'][2]['subworks'][19];
						$subwork25 = $g['retval'][2]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][2]['subworks'][21]);
						$dsr25 = $dsr25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][2]['subworks'][22]);
						$der25= $der25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop25 = $g['retval'][2]['subworks'][23];
						$obs25 = $g['retval'][2]['subworks'][24];
						$subwork26 = $g['retval'][2]['subworks'][25];
						$dsr26 = new DateTime($g['retval'][2]['subworks'][26]);
						$dsr26 = $dsr26->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der26 = new DateTime($g['retval'][2]['subworks'][27]);
						$der26= $der26->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2W 1ero 6Sub
					//2 W, 6 y 6
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 30 && count($g['retval'][1]['subworks']) == 30) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];
						$subwork15 = $g['retval'][0]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][0]['subworks'][21]);
						$dsr15 = $dsr15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][0]['subworks'][22]);
						$der15= $der15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][0]['subworks'][23];
						$obs15 = $g['retval'][0]['subworks'][24];
						$subwork16 = $g['retval'][0]['subworks'][25];
						$dsr16 = new DateTime($g['retval'][0]['subworks'][26]);
						$dsr16 = $dsr16->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der16 = new DateTime($g['retval'][0]['subworks'][27]);
						$der16= $der16->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop16 = $g['retval'][0]['subworks'][28];
						$obs16 = $g['retval'][0]['subworks'][29];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24= $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];
						$subwork25 = $g['retval'][1]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr25 = $dsr25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][1]['subworks'][22]);
						$der25= $der25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop25 = $g['retval'][1]['subworks'][23];
						$obs25 = $g['retval'][1]['subworks'][24];
						$subwork26 = $g['retval'][1]['subworks'][25];
						$dsr26 = new DateTime($g['retval'][1]['subworks'][26]);
						$dsr26 = $dsr26->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der26 = new DateTime($g['retval'][1]['subworks'][27]);
						$der26= $der26->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop26 = $g['retval'][1]['subworks'][28];
						$obs26 = $g['retval'][1]['subworks'][29];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}


					//DESCENDENTE
					//2 W, 2 y 1
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 10 && count($g['retval'][1]['subworks']) == 5) {

						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2W, 3 y 1
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 15 && count($g['retval'][1]['subworks']) == 5) {

						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2W, 4 y 1*** lo q falta en asc
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 20 && count($g['retval'][1]['subworks']) == 5) {

						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2W, 5 y 1
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 25 && count($g['retval'][1]['subworks']) == 5) {

						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];
						$subwork15 = $g['retval'][0]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][0]['subworks'][21]);
						$dsr15 = $dsr15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][0]['subworks'][22]);
						$der15= $der15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][0]['subworks'][23];
						$obs15 = $g['retval'][0]['subworks'][24];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2W, 6 y 1
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 30 && count($g['retval'][1]['subworks']) == 5) {

						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];
						$subwork15 = $g['retval'][0]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][0]['subworks'][21]);
						$dsr15 = $dsr15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][0]['subworks'][22]);
						$der15= $der15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][0]['subworks'][23];
						$obs15 = $g['retval'][0]['subworks'][24];
						$subwork16 = $g['retval'][0]['subworks'][25];
						$dsr16 = new DateTime($g['retval'][0]['subworks'][26]);
						$dsr16 = $dsr16->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der16 = new DateTime($g['retval'][0]['subworks'][27]);
						$der16= $der16->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop16 = $g['retval'][0]['subworks'][28];
						$obs16 = $g['retval'][0]['subworks'][29];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2W, 3 y 2  2 y 2 no es needed ya que tienen la misma cantidad de subworks
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 15 && count($g['retval'][1]['subworks']) == 10) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 4 y 2*
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 20 && count($g['retval'][1]['subworks']) == 10) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 5 y 2
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 25 && count($g['retval'][1]['subworks']) == 10) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];
						$subwork15 = $g['retval'][0]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][0]['subworks'][21]);
						$dsr15 = $dsr15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][0]['subworks'][22]);
						$der15= $der15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][0]['subworks'][23];
						$obs15 = $g['retval'][0]['subworks'][24];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 6 y 2
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 30 && count($g['retval'][1]['subworks']) == 10) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];
						$subwork15 = $g['retval'][0]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][0]['subworks'][21]);
						$dsr15 = $dsr15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][0]['subworks'][22]);
						$der15= $der15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][0]['subworks'][23];
						$obs15 = $g['retval'][0]['subworks'][24];
						$subwork16 = $g['retval'][0]['subworks'][25];
						$dsr16 = new DateTime($g['retval'][0]['subworks'][26]);
						$dsr16 = $dsr16->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der16 = new DateTime($g['retval'][0]['subworks'][27]);
						$der16= $der16->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop16 = $g['retval'][0]['subworks'][28];
						$obs16 = $g['retval'][0]['subworks'][29];


						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 4 y 3*
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 20 && count($g['retval'][1]['subworks']) == 15) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 5 y 3
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 25 && count($g['retval'][1]['subworks']) == 15) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];
						$subwork15 = $g['retval'][0]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][0]['subworks'][21]);
						$dsr15 = $dsr15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][0]['subworks'][22]);
						$der15= $der15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][0]['subworks'][23];
						$obs15 = $g['retval'][0]['subworks'][24];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 6 y 3
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 30 && count($g['retval'][1]['subworks']) == 15) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];
						$subwork15 = $g['retval'][0]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][0]['subworks'][21]);
						$dsr15 = $dsr15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][0]['subworks'][22]);
						$der15= $der15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][0]['subworks'][23];
						$obs15 = $g['retval'][0]['subworks'][24];
						$subwork16 = $g['retval'][0]['subworks'][25];
						$dsr16 = new DateTime($g['retval'][0]['subworks'][26]);
						$dsr16 = $dsr16->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der16 = new DateTime($g['retval'][0]['subworks'][27]);
						$der16= $der16->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop16 = $g['retval'][0]['subworks'][28];
						$obs16 = $g['retval'][0]['subworks'][29];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 5 y 4
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 25 && count($g['retval'][1]['subworks']) == 20) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];
						$subwork15 = $g['retval'][0]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][0]['subworks'][21]);
						$dsr15 = $dsr15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][0]['subworks'][22]);
						$der15= $der15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][0]['subworks'][23];
						$obs15 = $g['retval'][0]['subworks'][24];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24= $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 6 y 4
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 30 && count($g['retval'][1]['subworks']) == 20) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];
						$subwork15 = $g['retval'][0]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][0]['subworks'][21]);
						$dsr15 = $dsr15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][0]['subworks'][22]);
						$der15= $der15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][0]['subworks'][23];
						$obs15 = $g['retval'][0]['subworks'][24];
						$subwork16 = $g['retval'][0]['subworks'][25];
						$dsr16 = new DateTime($g['retval'][0]['subworks'][26]);
						$dsr16 = $dsr16->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der16 = new DateTime($g['retval'][0]['subworks'][27]);
						$der16= $der16->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop16 = $g['retval'][0]['subworks'][28];
						$obs16 = $g['retval'][0]['subworks'][29];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24= $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];


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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//2 W, 6 y 5
					if (count($g['retval']) == 2 && count($g['retval'][0]['subworks'])  == 30 && count($g['retval'][1]['subworks']) == 25) {
						$work1 = $g['retval'][0]['work'];
						$subwork11 = $g['retval'][0]['subworks'][0];
						$dsr11 = new DateTime($g['retval'][0]['subworks'][1]);
						$dsr11 = $dsr11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der11 = new DateTime($g['retval'][0]['subworks'][2]);
						$der11= $der11->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop11 = $g['retval'][0]['subworks'][3];
						$obs11 = $g['retval'][0]['subworks'][4];
						$subwork12 = $g['retval'][0]['subworks'][5];
						$dsr12 = new DateTime($g['retval'][0]['subworks'][6]);
						$dsr12 = $dsr12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der12 = new DateTime($g['retval'][0]['subworks'][7]);
						$der12= $der12->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop12 = $g['retval'][0]['subworks'][8];
						$obs12 = $g['retval'][0]['subworks'][9];
						$subwork13 = $g['retval'][0]['subworks'][10];
						$dsr13 = new DateTime($g['retval'][0]['subworks'][11]);
						$dsr13 = $dsr13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der13 = new DateTime($g['retval'][0]['subworks'][12]);
						$der13= $der13->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop13 = $g['retval'][0]['subworks'][13];
						$obs13 = $g['retval'][0]['subworks'][14];
						$subwork14 = $g['retval'][0]['subworks'][15];
						$dsr14 = new DateTime($g['retval'][0]['subworks'][16]);
						$dsr14 = $dsr14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der14 = new DateTime($g['retval'][0]['subworks'][17]);
						$der14= $der14->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop14 = $g['retval'][0]['subworks'][18];
						$obs14 = $g['retval'][0]['subworks'][19];
						$subwork15 = $g['retval'][0]['subworks'][20];
						$dsr15 = new DateTime($g['retval'][0]['subworks'][21]);
						$dsr15 = $dsr15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der15 = new DateTime($g['retval'][0]['subworks'][22]);
						$der15= $der15->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop15 = $g['retval'][0]['subworks'][23];
						$obs15 = $g['retval'][0]['subworks'][24];
						$subwork16 = $g['retval'][0]['subworks'][25];
						$dsr16 = new DateTime($g['retval'][0]['subworks'][26]);
						$dsr16 = $dsr16->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der16 = new DateTime($g['retval'][0]['subworks'][27]);
						$der16= $der16->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop16 = $g['retval'][0]['subworks'][28];
						$obs16 = $g['retval'][0]['subworks'][29];

						$work2 = $g['retval'][1]['work'];
						$subwork21 = $g['retval'][1]['subworks'][0];
						$dsr21 = new DateTime($g['retval'][1]['subworks'][1]);
						$dsr21 = $dsr21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der21 = new DateTime($g['retval'][1]['subworks'][2]);
						$der21= $der21->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop21 = $g['retval'][1]['subworks'][3];
						$obs21 = $g['retval'][1]['subworks'][4];
						$subwork22 = $g['retval'][1]['subworks'][5];
						$dsr22 = new DateTime($g['retval'][1]['subworks'][6]);
						$dsr22 = $dsr22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der22 = new DateTime($g['retval'][1]['subworks'][7]);
						$der22= $der22->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop22 = $g['retval'][1]['subworks'][8];
						$obs22 = $g['retval'][1]['subworks'][9];
						$subwork23 = $g['retval'][1]['subworks'][10];
						$dsr23 = new DateTime($g['retval'][1]['subworks'][11]);
						$dsr23 = $dsr23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der23 = new DateTime($g['retval'][1]['subworks'][12]);
						$der23= $der23->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop23 = $g['retval'][1]['subworks'][13];
						$obs23 = $g['retval'][1]['subworks'][14];
						$subwork24 = $g['retval'][1]['subworks'][15];
						$dsr24 = new DateTime($g['retval'][1]['subworks'][16]);
						$dsr24 = $dsr24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der24 = new DateTime($g['retval'][1]['subworks'][17]);
						$der24= $der24->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop24 = $g['retval'][1]['subworks'][18];
						$obs24 = $g['retval'][1]['subworks'][19];
						$subwork25 = $g['retval'][1]['subworks'][20];
						$dsr25 = new DateTime($g['retval'][1]['subworks'][21]);
						$dsr25 = $dsr25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$der25 = new DateTime($g['retval'][1]['subworks'][22]);
						$der25= $der25->setTimezone(new DateTimeZone('America/Santiago'))->format('d-m-Y, g:i a');
						$poop25 = $g['retval'][1]['subworks'][23];
						$obs25 = $g['retval'][1]['subworks'][24];

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

						$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
						$objWriter->save("ReportOut.xlsx");
						return View::make('DataSend.report', array("docRepor" => $docRepor));
					}
					//echo("Se ha superado max numero de trabajos");
					return Redirect::to('/dataform')->with('mensaje_error', 'Se ha superado max numero de trabajos');




				}//else



			}//first if

			//return View::make('DataSend.report', array("docRepor" => $docRepor));
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
		//recibo los datos de APP movil
		$aRequest = json_decode(file_get_contents('php://input'),true);

		/*try {
			$fichero=fopen('test.log','w');
		} catch (Exception $e) {
			echo "capturada";
			//chmod('test.log', 0777);
		}
		$fichero=fopen('test.log','w');
	 		if($fichero == false) {
   			die("No se ha podido crear el archivo.");
		}
		fwrite($fichero,json_encode($aRequest));
		fclose($fichero);*/

		//guardo nombre de Equipos
		$m = new MongoClient();//obsoleta desde mongo 1.0.0
		$db = $m->SenditForm;

		$equipments = array("equi" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['EQUIPMENT']);
		$coll_equipments = $db->equipments;
		if($coll_equipments->count() == 0){
			$coll_equipments->insert($equipments);
		}else{
			$result = $coll_equipments->findOne(["equi" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['EQUIPMENT']]);
			if (!$result) {
			$coll_equipments->insert($equipments);
			}
		}
		//guardo ubicaciones
		$locs = array("loc" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['LOCALIZATION_EQUIPMENT']);
		$coll_locs = $db->locs;
		if($coll_locs->count() == 0){
			$coll_locs->insert($locs);
		}else{
			$result = $coll_locs->findOne(["loc" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['LOCALIZATION_EQUIPMENT']]);
			if (!$result) {
			$coll_locs->insert($locs);
			}
		}
		//guardo identificaciones
		$idens = array("iden" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['IDENTIFICATION_EQUIPMENT']);
		$coll_idens = $db->idens;
		if($coll_idens->count() == 0){
			$coll_idens->insert($idens);
		}else{
			$result = $coll_idens->findOne(["iden" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['IDENTIFICATION_EQUIPMENT']]);
			if (!$result) {
			$coll_idens->insert($idens);
			}
		}

		//guardo datos que vienen del app movil
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
						"HOUR_PROG" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['HOUR_PROG'],
						"WORK" =>  array(
							"WORK_NUEVO" => "SI",
							"WORK_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['WORK'],
							"SUBWORK" => array(
								"SUBWORK_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['SUBWORK'],
								"DATE_START_REAL" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_REAL'],
								"DATE_END_REAL" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_REAL'],
								"POOP" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['POOP'],
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
						"HOUR_PROG" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['HOUR_PROG'],
						"WORK" =>  array(
							"WORK_NUEVO" => "NO",
							"WORK_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['WORK'],
							"SUBWORK" => array(
								"SUBWORK_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['SUBWORK'],
								"DATE_START_REAL" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_REAL'],
								"DATE_END_REAL" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_REAL'],
								"POOP" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['POOP'],
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
			//verifico que no se inserte una mismo subtrabajo con la misma fecha de programacion
			$result = $collRepor->findOne([
				'EQUIPMENT.WORK.WORK_NAME' => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['WORK'],
				'EQUIPMENT.WORK.SUBWORK.SUBWORK_NAME' => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['SUBWORK'],
				'EQUIPMENT.DATE_START_PROGRAMMED' => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_PROGRAMMED'],
				'EQUIPMENT.DATE_END_PROGRAMMED' => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_PROGRAMMED'],
				'EQUIPMENT.EQUIPMENT_NAME' => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['EQUIPMENT'],
				/*'EQUIPMENT.LOCALIZATION_EQUIPMENT.LOCALIZATION_NAME' => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['LOCALIZATION_EQUIPMENT'],
				'EQUIPMENT.IDENTIFICATION_EQUIPMENT.IDENTIFICATION_NAME' => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['IDENTIFICATION_EQUIPMENT']*/
				]);
			if (!$result) {

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
						"HOUR_PROG" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['HOUR_PROG'],
						"WORK" =>  array(
							"WORK_NUEVO" => "SI",
							"WORK_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['WORK'],
							"SUBWORK" => array(
								"SUBWORK_NAME" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['SUBWORK'],
								"DATE_START_REAL" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_REAL'],
								"DATE_END_REAL" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_REAL'],
								"POOP" => $aRequest['Entry']['AnswersJson']['ADD_WORK_PAGE']['POOP'],
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
				$coll_same_sub = $db->Same_subw;
				$docRepor = $coll_same_sub->insert($array);
				echo "Insertado en collection Same_subw";
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
