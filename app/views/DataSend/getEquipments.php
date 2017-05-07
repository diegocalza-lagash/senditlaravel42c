<?php
	if (isset($_GET['equipment'])) {
			$equipment = ($_GET['equipment']);
			$m = new MongoClient();
			$db = $m->SenditForm;
			$collwf = $db->Repor;
			$equipos = $collwf->find();
			foreach ($equipos as $v) {
				$equipo = $v['EQUIPMENT']['EQUIPMENT_NAME'];
			}
			switch ($equipment) {
				case 'c':
					$e['nombre'] = "Caldera";
					break;

				case 'ca':

					break;
			}

		}
		echo json_encode($e);
?>