@extends("layouts.master")
@section('title','Lista De Trabajos')
@section('sidebar')
@parent
<li><a href="/download">Exportar</a></li>
@stop
@section('trabajos')
@parent
<li><a href="/dataform">Ver Todos Los Trabajos</a></li>
@stop
@section('content')
	<div class="dataTable_wrapper">
	<h1 class="sub-header">Lista De Trabajos</h1>
		<div class="dataTable_header">
			<div class = "dataTable_download">
			<!--<a href="/download">Exportar</a>-->
			</div>
			<div class="dataTable_filtro"></div>
		</div>
		<table class="table table-striped table-condensed table-hover table-bordered">
			<thead>
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
				</tr>
			</thead>
			<tbody>
				@foreach ($docRepor as $row)
					<tr>

						<td><?php
						$startTime = new DateTime($row['Entry']['StartTime']);
						//$startTime->setTimezone(new DateTimeZone('America/Santiago'));
						echo $startTime->format('j F, Y, g:i a');
							?>
							<div >
								<span><b>Subido:</b></span>
							</div>
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
						echo $DATE_START_REAL->format('d-m-Y, g:i a');
						?></td>
						<td><?php
						$DATE_END_REAL = new DateTime($row['EQUIPMENT']['WORK']['SUBWORK']['DATE_END_REAL']);
						$DATE_END_REAL->setTimezone(new DateTimeZone('America/Santiago'));
						echo $DATE_END_REAL->format('d-m-Y, g:i a');
						?></td>
						<td><?php echo $row['EQUIPMENT']['WORK']['SUBWORK']['POOP']."%"?></td>
						<td><?php echo $row['EQUIPMENT']['WORK']['SUBWORK']['OBSERVATIONS']?></td>
						<td>
							<?php
							$id = $row['Entry']['Id'];
							$Id = substr($id, 0, 8).'-'.substr($id, 8, 4).'-'.substr($id, 12, 4).'-'.substr($id, 16, 4).'-'.substr($id, 20, 32);
								echo '<a href="https://app.sendit.cl/Files/FormEntry/'.$row['ProviderId'].'-'.$Id.$row['EQUIPMENT']['WORK']['PHOTOS']['PHOTO1'].'">Ver Foto</a>'
							?>
						</td>


					</tr>
				@endforeach
			</tbody>
		</table>
	</div>
@stop

