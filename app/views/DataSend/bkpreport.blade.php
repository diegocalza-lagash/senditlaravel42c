@extends("layouts.master")
@section('title','Lista De Trabajos')
@section('content')
	<div class="box-body">
		<div id ="DataTables_Table_0_wrapper" class="dataTables_wrapper no-footer">
			<!--<div class="dataTables_actions"></div>-->
			<div id="DataTables_0_Filter" class="dataTables_filter">
				<!--<button id="showSideFilter" style="display: none; " class="btn form-control btn-primary pull-right input-sm" type="button"><i class="fa fa-filter"></i></button><label><input type="search" placeholder="Find answer that starts with..." class="form-control input-sm" aria-controls="DataTables_Table_0"></label>-->

			</div>
			<!--<div id="DataTables_0_processing" class="" ><span class="dataTables_processing panel panel-default"></span>
			</div>-->
			<h2 class="sub-header">Lista De Trabajos</h2>
			<a href="/download">Descargar Contenido</a></td>
			<div class="dataTables_scroll">
				<div class="dataTables_scrollHead">
					<div class="dataTables_scrollHeadInner">
						<table class="display nowrap dataTable no-footer">
							<thead class="">
								<tr >
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
									<!--<th>Descripción Foto 1</th>-->
									<!--<th>Foto 2</th>
									<!--<th>Descripción Foto 2</th>
									<th>Foto 3</th>
									<!--<th>Descripción Foto 3</th>
									<th>Video</th>
									<!--<th>Descripción Video</th>-->
								</tr>
							</thead>
						</table>
					</div>
					<div></div>
				</div>
				<div class="dataTables_scrollBody">
					<table id="DataTable" class="display nowrap dataTable no-footer">
						<thead class="">
							<tr>

							</tr>
						</thead>
							<tbody>
								@foreach ($docRepor as $row)
									<tr>

										<td><?php
										$startTime = new DateTime($row['Entry']['StartTime']);
										$startTime->setTimezone(new DateTimeZone('America/Santiago'));
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
										echo $DATE_START_REAL->format('j F, Y, g:i a');
										?></td>
										<td><?php
										$DATE_END_REAL = new DateTime($row['EQUIPMENT']['WORK']['SUBWORK']['DATE_END_REAL']);
										$DATE_END_REAL->setTimezone(new DateTimeZone('America/Santiago'));
										echo $DATE_START_REAL->format('j F, Y, g:i a');
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
					<div id ="DataTables_Table_0_length" class="dataTables_length">
						<label> Mostrar
						<select class="" name="" aria-controls="DataTables_Table_0" class="form-control input-sm"><option value="10">10</option><option value="25">25</option><option value="50">50</option><option value="100">100</option></select>
						</label>
					</div>
				</div>

			</div>

		</div>

	</div>

@stop
<script type="text/javascript">
	/*$(document).ready(function() {
    $('#DataTable').DataTable( {
        "scrollX": true
    } );
} );*/
</script>
