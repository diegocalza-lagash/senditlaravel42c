@extends("layouts.master")
@section('title','Todos Los Trabajos')
@section('content')
<?php
	$m = new MongoClient();//obsoleta desde mongo 1.0.0
	$db = $m->SenditForm;
	$collRepor = $db->Repor;
	$docRepor = $collRepor->find();
?>
@if(Session::has('mensaje_error'))
    <div class="alert alert-danger">{{ Session::get('mensaje_error') }}</div>
@endif
<!--
Form::macro('myField', function()
{
    return '<input type="awesome">';
});
-->
<!--<script type="text/javascript">
	input date para mozilla
	$(function() {
     $( "#input_date" ).datepicker({ dateFormat: 'yy-mm-dd'});
});-->
</script>
<div class="dataTable_wrapper">
	<div class="dataTable_form">
		{{ Form::open(array('url' => 'report/show','method' => 'get')) }}
	    {{ Form::label('equipo','Equipo',['required' => 'true']) }}
	    {{ Form::text('equi','Caldera', $attributes = array('placeholder'=>"Caldera","id" =>"equipo",'required' => 'true')) }}
	    {{ Form::label('loc','Ubicación') }}
	    {{ Form::text('loc','Economizador II piso 6°, Buzón Eco 2',['required' => 'true']) }}
	    {{ Form::label('iden','Identificación') }}
	    {{ Form::text('iden','Poder',['required' => 'true'])}}
	    {{ Form::label('dsp','FIP') }}
	    {{Form::input('date', 'dsp', null, ['class' => '', 'placeholder' => 'dd/mm/yyyy','id' => 'input_date','required' => 'true']) }}
	    {{ Form::label('dep','FEP') }}
	    {{Form::input('date', 'dep', null, ['class' => '', 'placeholder' => 'dd/mm/yyyy','id' => 'input_date','required' => 'true']) }}

	    <!--{{ Form::text('dep','FTP') }}-->
	    {{ Form::submit('Buscar'); }}
	{{ Form::close() }}
	<!--<form class="ng-pristine ng-valid" method="GET" action="report/show" >
	    <label Equipo</label>
	    <input placeholder="Caldera" id="equipo" name="equi" value="Caldera" type="text">
	    <label>Ubicación</label>
	    <input name="loc" value="Economizador II piso 6°, Buzón Eco 2" id="loc" type="text">
	    <label>Identificación</label>
	    <input name="iden" value="Poder" id="iden" type="text">
	   <label >FIP</label>
	    <input name="dep" value="15/04/2017" type="month" required>
	    <label >FEP</label>
	     <input name="dep" value="15/04/2017" type="date">
	    <input value="Buscar" type="submit">

	</form>-->
	</div>
	<script type="text/javascript">
	/*$(document).ready(function(){
		$("#equipo").click(function(){

			$equipment = $("#equipo").val();
			if ($equipment != "") {
					$.ajax({
					type 		:"get",
					url 		:"DataSend/getEquipments.php",
					data 		:{equipment: $(this).val()},
					dataType 	:"json",
					success 	:function(data){
						$(".hint ul ").append("<li>data.nombre</li>");
						console.log(data.nombre);
					}
				});
			}

		})
	})*/
	</script>
	<div class="hint">
		<ul>

		</ul>
	</div>
	<table id= "lista-crud" class="table table-striped table-hover table-bordered table-condensed listar-act">
		<thead>
			<tr>
				<th>Fecha De Envío</th>
				<th>Enviado por</th>
				<th>Ubicación</th>
				<th>Equipo</th>
				<th>Identificación Equipo</th>
				<th>Sistema de bloqueo</th>
				<th>Trabajo</th>
				<th>SubTrabajo</th>
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
					<td><?php echo $row['EQUIPMENT']['LOCALIZATION_EQUIPMENT']['LOCALIZATION_NAME']?></td>
					<td><?php echo $row['EQUIPMENT']['EQUIPMENT_NAME']?></td>
					<td><?php echo $row['EQUIPMENT']['IDENTIFICATION_EQUIPMENT']['IDENTIFICATION_NAME']?></td>
					<td><?php echo $row['EQUIPMENT']['BLOCK_SYSTEM']?></td>
					<td><?php echo $row['EQUIPMENT']['WORK']['WORK_NAME']?></td>
					<td><?php echo $row['EQUIPMENT']['WORK']['SUBWORK']['SUBWORK_NAME']?></td>
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
				<?php
			}
				?>
		</tbody>
	</table>
@stop
</div>



