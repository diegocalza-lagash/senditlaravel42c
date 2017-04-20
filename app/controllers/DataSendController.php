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

				if (!$docRepor -> count()) {
					echo "Sin Trabajos";
				}else{
					foreach ($docRepor as $arr) {
						$work_name = $arr['EQUIPMENT']['WORK']['WORK_NAME'];
						$subwork_name = $arr['EQUIPMENT']['WORK']['SUBWORK']['SUBWORK_NAME'];
						$subwork_dsr = $arr['EQUIPMENT']['WORK']['SUBWORK']['DATE_START_REAL'];
						$subwork_der = $arr['EQUIPMENT']['WORK']['SUBWORK']['DATE_END_REAL'];
						$subwork_poop = $arr['EQUIPMENT']['WORK']['SUBWORK']['POOP'];
						$subwork_obs = $arr['EQUIPMENT']['WORK']['SUBWORK']['OBSERVATIONS'];
						$dsp = $arr['EQUIPMENT']['DATE_START_PROGRAMMED'];
						$dep = $arr['EQUIPMENT']['DATE_END_PROGRAMMED'];

						echo '<pre>'.var_dump($work_name,$subwork_name,$subwork_poop,$dsp,$dep).'</pre>';
						$work = new Work;
						$work->nombre = $work_name;
						$subwork = new Subwork;
						$subwork->nombre = $subwork_name;
						$subwork->fecha_inicio_real = $subwork_dsr;
						$subwork->fecha_termino_real = $subwork_der;
						$subwork->poop = $subwork_poop;
						$subwork->observaciones = $subwork_obs;
					}
					if ($work->save() && $subwork->save()) {
							//echo "insertado en work model and sub_work model";
							$first = DB::table('works')->distinct()
									->join('subworks','works.id','=','subworks.id')
									->select('subworks.nombre');

							$query = DB::table('works')->distinct()
									->join('subworks','works.id','=','subworks.id')
									->select('works.nombre')
									->union($first)
									->get();
							foreach ($query as $q) {
								var_dump($q->nombre);
							}


						}


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
						$obs = $docRepor['EQUIPMENT']['WORK']['SUBWORK']['OBSERVATIONS'];*/

					}


					//return View::make('DataSend.report',array("docRepor" => $docRepor));

				}

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

		$m = new MongoClient();//obsoleta desde mongo 1.0.0
		$db = $m->SenditForm;
		$collRepor = $db->Repor;
		$docRepor = $collRepor->insert($array);
		echo "Insertado en Repor collec";
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


	}

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
