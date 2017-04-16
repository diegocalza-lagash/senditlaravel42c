<?php
		$m = new MongoClient();
		$db = $m->SenditForm;
		$collWorks = $db->Works;
		$docsWorks = $collWorks->find();

?>
<!DOCTYPE html>
<html>
<head>
	<title>hola kalza</title>
	<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css" integrity="sha384-1q8mTJOASx8j1Au+a5WDVnPi2lkFfwwEAa8hDDdjZlpLegxhjVME1fgjWPGmkzs7" crossorigin="anonymous">
</head>
<body>
	<table id= "lista-crud" class="table table-striped table-condensed listar-act">
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

					<td><?php
					$startTime = new DateTime($row['Entry']['StartTime']);
					$startTime->setTimezone(new DateTimeZone('America/Santiago'));
					echo $startTime->format('j F, Y, g:i a');
						?>
					</td>
					<td><?php echo $row['Entry']['UserFirstName'].$row['Entry']['UserLastName']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['ADD_WORK_PAGE']['WORK']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['ADD_WORK_PAGE']['SUBWORK']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['ADD_WORK_PAGE']['BLOCK_SYSTEM']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_START_PROGRAMMED']?></td>
					<td><?php echo $row['Entry']['AnswersJson']['ADD_WORK_PAGE']['DATE_END_PROGRAMMED']?></td>

					<!--<td>{{ HTML::linkAction('DataSendController@report','Descargar Informe') }}</td>-->
					<td><a href="excel/{{$row['Entry']['Id']}}">Descargar Contenido</a></td>
				</tr>
				<?php
			}
				?>
		</tbody>
	</table>
</body>
</html>