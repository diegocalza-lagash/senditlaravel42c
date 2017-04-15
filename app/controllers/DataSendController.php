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
		//
		//echo "hola";
		/*$m = new MongoDB\Client();
		$db = $m->formSendit2;
		$collection = $db->DataFormTest;
		$docSendit = $collection->find();
		foreach ($docSendit as $row) {
			# code...
			//print_r($row);
			echo $row->Entry->UserEmail;
		}*/
		return View::make('dataSends.index');

	}
	public function report($id){


		//echo "hola report: ".$id;
		$m = new MongoDB\Client();
		$db = $m->formSendit2;
		$collection = $db->DataFormTest;
		//$docSendit = $collection->find(['ProviderId' => $id]);
		//$db->users->find(array("age" => array('$gt' => 33, '$lte' => 40)));
		$docSendit = $collection->findOne(['Entry.Id' => $id]);

		//echo "pase";
		//print_r($docSendit);
		//var_dump($docSendit);
		/*echo $docSendit['Entry']['StartTime'];
		echo $docSendit['Entry']['UserFirstName'].$docSendit['Entry']['UserLastName'];
		echo $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['mantencion_equipos'];
		echo $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['Trabajos'];
		echo $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['Sub_trabajos'];
		echo $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['Sistema_bloqueo'];
		echo $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_inicio_prog'];
		echo $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_termino_prog'];
		echo $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_inicio_real'];
		echo $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_termino_real'];
		echo $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['porcentaje_avance_fisico'];
		echo $docSendit['Entry']['AnswersJson']['Trabajos_planificados2']['observaciones'];*/

		/* $StartTime = $docSendit['Entry']['StartTime'];
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

		Excel::create("report2",function($excel){

			$excel->sheet('kalza',function($sheet){
				//$UserFirstName =$docSendit['Entry']['UserFirstName'];
				$data=[];
				array_push($data, array(
							array('data1', 'data2'),
					    	array('data3', 'data4')));
				/*$data = array(
					    	array('data1', 'data2'),
					    	array('data3', 'data4')
							);

				$sheet->with($data);

			});

		})->export('xls');*/

		//return View::make('dataSends.index');
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
	//echo "holapost"." ";
		//ob_start();
		$aRequest = json_decode(file_get_contents('php://input'),true);
		//print_r($aRequest);
		//echo $aRequest['Entry']['UserEmail'];
		/*foreach($aRequest as $obj){
	        $email = $obj->Entry->UserEmail;
	        //$mantencion_equipos = $obj->AnswersJson->Trabajos_planificados2->mantencion_equipos;
	        echo $email;//." ".$mantencion_equipos." ";
		}*/
		//echo "hola again";
		$fichero=fopen('test.log','w');
	 		if($fichero == false) {
   			die("No se ha podido crear el archivo.");
		}
		fwrite($fichero,json_encode($aRequest));
		fclose($fichero);
//require 'vendor/autoload.php';




		//$m = new MongoDB\Driver\Manager("mongodb://localhost:27017");
		$m = new MongoClient();//obsoleta desde mongo 1.0.0

		//
	//	$m = new MongoDB\Client();
		//$m = new MongoDB\Client();


		$db = $m->formSendit2;

		$db = $m->SenditForm;


		$collWorks = $db->Works;
		//var_dump($collection);
		//require 'vendor/autoload.php';

		 $doc = $collection->insert($aRequest);

		echo "doc insertado";
		 echo "doc insertado";


		$Work = $aRequest['Entry']['AnswersJson']['Trabajos_planificados2']['Trabajos'];
		$Sub_W = $aRequest['Entry']['AnswersJson']['Trabajos_planificados2']['Sub_trabajos'];
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
		}*/
		$docWork = $collWorks->findOne(['Entry.AnswersJson.Trabajos_planificados2.Trabajos' => $Work]);
		//var_dump($docWork);
		$work = $docWork['Entry']['AnswersJson']['Trabajos_planificados2']['Trabajos'];
		//echo $work;
		if ($docWork){
			# code...
			//$docSendit = $collection->find(['Entry.AnswersJson.Trabajos_planificados2.Trabajos' => $Work]);
			$IdForm = $docWork['Entry']['Id'];//get id de Works collection
			$collSubWorks = $db->SubWorks;//create collection

			$docSubWs = $collSubWorks->insert($aRequest);
			//$subws = $collSubWorks->find();
			$updateResult = $collSubWorks->update(
				['Entry.AnswersJson.Trabajos_planificados2.Trabajos' => $work],
				[ '$set' => ['Entry.Id' => $IdForm]],
				['multiple' => true]
			);
			/*foreach ($subws as $subw) {
				$updateResult = $subw->update(
			    ['Entry.AnswersJson.Trabajos_planificados2.Trabajos' => $work],
			    [ '$set' => ['Entry.Id' => $IdForm]]
			);
			$work = $subw['Entry']['AnswersJson']['Trabajos_planificados2']['Trabajos'];
			echo $work;
			//echo $subW->Entry->AnswersJson->Trabajos_planificados2->Trabajos;

			}
			/*for($i=0;$i<count($subws);$i++){
				$id_fruta=$subws[$i]->Entry->AnswersJson->Trabajos_planificados2->Trabajos;

			    echo $id_fruta;
			}*/

			echo "doc insertado en SubWs collection";
		}else{

		$docWork = $collWorks->insert($aRequest);
		//$indexName = $collection->createIndex(['borough' => 1, 'cuisine' => 1]);
		echo "doc insertado en Works collection";

		}


		//$doc = $collection->insert($aRequest);
		//echo "doc insertado";

		 //$email = $aRequest['Entry']['UserEmail'];


		 $email = $aRequest['Entry']['UserEmail'];
		echo $email;


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
