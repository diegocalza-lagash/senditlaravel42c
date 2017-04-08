<?php
		$m = new MongoDB\Client();
		$db = $m->formSendit2;
		$collection = $db->DataFormTest;
		$docSendit = $collection->find();
		//foreach ($docSendit as $row) {
			# code...
			//print_r($docSendit);
			//echo $row->Entry->UserEmail;
	//	}
?>
<!DOCTYPE html>
<html>
<head>
	<title>hola kalza</title>
</head>
<body>
	<table>
		<thead>
			<tr>


				<th>Fecha De Envío</th>
				<th>Enviado por</th>
				<!--<th>Ubicación</th>-->
				<th>Mantencion de Equipos</th>
				<th>Trabajos</th>
				<th>SubTrabajos</th>
				<th>Sistema de bloqueo</th>
				<th>Fecha De Inicio Programada</th>
				<th>Fecha De Término Programada</th>
				<th>Fecha De Inicio Real</th>
				<th>Fecha De Término Real</th>
				<th>Porcentaje De Avance Físico</th>
				<th>Observaciones</th>
			</tr>
		</thead>
		<tbody>
			<?php
			foreach ($docSendit as $row) {
				?>
				<tr>

					<td><?php echo $row['Entry']['StartTime']?></td>
					<td><?php echo $row['Entry']['UserFirstName'].$row['Entry']['UserLastName']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['Trabajos_planificados2']['mantencion_equipos']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['Trabajos_planificados2']['Trabajos']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['Trabajos_planificados2']['Sub_trabajos']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['Trabajos_planificados2']['Sistema_bloqueo']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_inicio_prog']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_termino_prog']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_inicio_real']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['Trabajos_planificados2']['fecha_termino_real']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['Trabajos_planificados2']['porcentaje_avance_fisico']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['Trabajos_planificados2']['observaciones']?></td>
					<!--<td>{{ HTML::linkAction('DataSendController@report','Descargar Informe') }}</td>-->
					<td><a href="excel/{{$row['Entry']['Id']}}">Descargar Informe</a></td>
				</tr>
				<?php
			}
				?>
		</tbody>
	</table>
</body>
</html>