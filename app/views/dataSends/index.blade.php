<?php
		$m = new MongoClient();
		$db = $m->SenditForm;
		$collWorks = $db->Works;
		$docsWorks = $collWorks->find();
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
				<th>Trabajos</th>
				<th>SubTrabajos</th>
				<th>Sistema de bloqueo</th>
				<th>Fecha De Inicio Programada</th>
				<th>Fecha De Término Programada</th>
				<!--<th>Fecha De Inicio Real</th>
				<th>Fecha De Término Real</th>-->

			</tr>
		</thead>
		<tbody>
			<?php
			foreach ($docsWorks as $row) {
				?>
				<tr>

					<td><?php echo $row['Entry']['StartTime']?></td>
					<td><?php echo $row['Entry']['UserFirstName'].$row['Entry']['UserLastName']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['ADD_WORK_PAGE']['WORK']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['ADD_WORK_PAGE']['SUBWORK']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['ADD_WORK_PAGE']['BLOCK_SYSTEM']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_PROGRAMMED']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_PROGRAMMED']?></td>

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