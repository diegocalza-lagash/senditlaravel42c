<!DOCTYPE html>
<html>
<head>
	<meta charset="utf-8">
	<title>hola kalza</title>
	<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css" integrity="sha384-1q8mTJOASx8j1Au+a5WDVnPi2lkFfwwEAa8hDDdjZlpLegxhjVME1fgjWPGmkzs7" crossorigin="anonymous">
	<link rel="stylesheet" href="{{ URL::asset('assets/css/css-table.css') }}">
</head>
<body>
	<a href="/download">Descargar Contenido</a></td>
	<div>
		<table  class="table table-striped  table-condensed  listar-act ">
		<thead >
			<tr>

				<th>Fecha De Envío</th>
				<th>Enviado por</th>
				<th>Trabajo</th>
				<th>SubTrabajo</th>
				<th>Ubicación</th>
				<th>Equipo</th>
				<th>Identificación Equipo</th>
				<th>Sistema de bloqueo</th>
				<th>Fecha De Inicio Programada</th>
				<th>Fecha De Término Programada</th>
				<th>Fecha De Inicio Real</th>
				<th>Fecha De Término Real</th>
				<th>Avance</th>
				<th>Observaciones</th>
				<th>Foto 1</th>
				<th>Descripción Foto 1</th>
				<th>Foto 2</th>
				<th>Descripción Foto 2</th>
				<th>Foto 3</th>
				<th>Descripción Foto 3</th>
				<th>Video</th>
				<th>Descripción Video</th>

			</tr>
		</thead>
		<tbody>
			<?php
			foreach ($docRepor as $row) {
				?>
				<tr>

					<td><?php
					$startTime = new DateTime($row['Entry']['StartTime']);
					$startTime->setTimezone(new DateTimeZone('America/Santiago'));
					echo $startTime->format('j F, Y, g:i a');
						?>
					</td>
					<td><?php echo $row['Entry']['UserFirstName']." ".$row['Entry']['UserLastName']?></td>
					<td><?php echo $row['EQUIPMENT']['WORK']['WORK_NAME']?></td>
					<td><?php echo $row['EQUIPMENT']['WORK']['SUBWORK']['SUBWORK_NAME']?></td>
					<td><?php echo $row['EQUIPMENT']['LOCALIZATION_EQUIPMENT']['LOCALIZATION_NAME']?></td>
					<td><?php echo $row['EQUIPMENT']['EQUIPMENT_NAME']?></td>
					<td><?php echo $row['EQUIPMENT']['IDENTIFICATION_EQUIPMENT']['IDENTIFICATION_NAME']?></td>
					<td><?php echo $row['EQUIPMENT']['BLOCK_SYSTEM']?></td>
					<td><?php echo $row['EQUIPMENT']['DATE_START_PROGRAMMED']?></td>
					<td><?php echo $row['EQUIPMENT']['DATE_END_PROGRAMMED']?></td>
					<td><?php
					$DATE_START_REAL = new DateTime($row['EQUIPMENT']['WORK']['SUBWORK']['DATE_START_REAL']);
					$DATE_START_REAL->setTimezone(new DateTimeZone('America/Santiago'));
					echo $DATE_START_REAL->format('j F, Y, g:i a');
					?></td>
					<td><?php
					$DATE_END_REAL = new DateTime($row['EQUIPMENT']['WORK']['SUBWORK']['DATE_END_REAL']);
					$DATE_END_REAL->setTimezone(new DateTimeZone('America/Santiago'));
					echo $DATE_START_REAL->format('j F, Y, g:i a');
					?></td>
					<td><?php echo $row['EQUIPMENT']['WORK']['SUBWORK']['POOP']."%"?></td>
					<td><?php echo $row['EQUIPMENT']['WORK']['SUBWORK']['OBSERVATIONS']?></td>
					<td><?php echo $row['EQUIPMENT']['WORK']['PHOTOS']['PHOTO1']?></td>
					<td><?php echo $row['EQUIPMENT']['WORK']['PHOTOS']['DESCRIPTION_PHOTO1']?></td>
					<td><?php echo $row['EQUIPMENT']['WORK']['PHOTOS']['PHOTO2']?></td>
					<td><?php echo $row['EQUIPMENT']['WORK']['PHOTOS']['DESCRIPTION_PHOTO2']?></td>
					<td><?php echo $row['EQUIPMENT']['WORK']['PHOTOS']['PHOTO3']?></td>
					<td><?php echo $row['EQUIPMENT']['WORK']['PHOTOS']['DESCRIPTION_PHOTO3']?></td>
					<td><?php echo $row['EQUIPMENT']['WORK']['PHOTOS']['VIDEO']?></td>
					<td><?php echo $row['EQUIPMENT']['WORK']['PHOTOS']['DESCRIPTION_VIDEO']?></td>

					<!--<td>{{ HTML::linkAction('DataSendController@report','Descargar Informe') }}</td>-->
				</tr>
				<?php
			}
				?>
		</tbody>
	</table>
	</div>

</body>
</html>
