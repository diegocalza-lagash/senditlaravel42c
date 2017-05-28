@extends("layouts.master")
@section('title','Todos Los Trabajos')
@section('content')
<?php
	$m = new MongoClient();//obsoleta desde mongo 1.0.0
	$db = $m->SenditForm;
	//obtengo todos los trabajos
	$collRepor = $db->Repor;
	$docRepor = $collRepor->find();
	//obtengo equipos
	$equipments = $db->equipments->find();
	//obtengos locs
	$locs =$db->locs->find();
	//obtengo identificaciones
	$idens =$db->idens->find();

?>

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
@if(Session::has('mensaje_error'))
<div class="alert alert-danger">{{ Session::get('mensaje_error') }}</div>
@endif
<div class="dataTable_form">


	    {{ Form::open(array('url' => 'report/show','method' => 'get','style' => '')) }}
	    	<label> Equipo </label>
	    	<select id ="equi_select">
		    @foreach($equipments as $e)
		    	<!--<option>@if(isset($e['equi'])) {{ $e['equi'] }} @endif</option>-->
		    	<option value="{{$e['equi']}}"> {{ $e['equi'] }} </option>
		    @endforeach
		    </select>

		    <label>Ubicación</label>
		     <select  id ="loc_select"  >
		    @foreach($locs as $l)
		    	<option value="{{ $l['loc'] }}">{{ $l['loc'] }}</option>
		    @endforeach
		    </select>

		    <label>Identificación</label>
		     <select id="iden_select" >
		    @foreach($idens as $i)
		    	<option value="{{ $i['iden']}}"> {{ $i['iden'] }} </option>
		    @endforeach
		    </select>
			    {{ Form::text('equi', null, array('placeholder'=>"Caldera","id" =>"equi_input",'required' => 'true', 'hidden' => 'true')) }}
			    {{ Form::text('loc', null, ['required' => 'true' ,"id" => "loc_input", 'hidden' => 'true' ]) }}
			    {{ Form::text('iden', null, ['required' => 'true', "id" => "iden_input", 'hidden' => 'true' ])}}
			    {{ Form::label('dsp','FIP') }}
			    {{ Form::input('date','dsp', null, ['class' => '', 'placeholder' => 'yyyy-mm-dd','id' => 'date_fip','required' => 'true']) }}
			    {{ Form::label('dep','FTP') }}
			    {{Form::input('date', 'dep', null, ['class' => '', 'placeholder' => 'yyyy-mm-dd','id' => 'date_ftp','required' => 'true']) }}

			    {{ Form::submit('Buscar'); }}
		{{ Form::close() }}
	</div>
<div class="dataTable_wrapper" style ="/*margin-left: 12%; margin-top: 3%;width: 87%;border: 1px solid #d9d9d9; padding: 5px;*/">


<script type="text/javascript">
	$(document).ready(function(){
		//Asignar valores a los inputs del form
		//asgina el valor por defecto del dropdrownlist
		if ($("#equi_select").val() != null && $("#loc_select").val() !=null && $("#iden_select") != null ) {

			$("#equi_input").attr("value", $("#equi_select").val());
			$("#loc_input").attr("value", $("#loc_select").val());
			$("#iden_input").attr("value",$("#iden_select").val());
			console.log("ok1");
			console.log("ok1");
		}
		//asigna el valor cuando cambien la lista desplegable
		$("#equi_select").change(function(){
			var equi = $(this).val();
			$("#equi_input").attr("value", equi);
			console.log("ok");
		})
		$("#loc_select").change(function(){
			$("#loc_input").attr("value", $(this).val());
		})
		$("#iden_select").change(function(){
			$("#iden_input").attr("value", $(this).val());
		})

		//validar formato de las fechas antes del envio
		$("form").submit(function(e){
			var fip = $("#date_fip").val();
			var ftp = $("#date_ftp").val();
			//console.log(fip,ftp);
				if(validarFormatoFecha(fip,ftp)){
					if (existeFecha(fip) && existeFecha(ftp)) {
						if(ftpMayorFip(fip,ftp)){
							//alert("fecha correcta");
							//e.preventDefault();
					    }else{
				            alert("FTP no puede ser menor a FIP");
				            e.preventDefault();
					    }
					}else{
						alert("la fecha no es del calendario");
						e.preventDefault();
					}

				      //alert("fecha correcta")
				}else{
				      alert("El formato de la fecha es incorrecto.");
				      e.preventDefault();
				}

		})//form

	})//document
	//VALIDO FORMATO DE LA FECHAS
	function validarFormatoFecha(fip,ftp) {
		//formato dd-mm-yyyy
	      var RegExPattern = /^\d{2,4}\-\d{1,2}\-\d{1,2}$/;///^\d{1,2}\/\d{1,2}\/\d{2,4}$/; -> formato dd/mm/yyyy
	      if ((fip.match(RegExPattern)) && (fip!='' && ftp.match(RegExPattern) && ftp!='')) {
	            return true;
	      } else {
	            return false;
	      }
	}
	//VALIDO QUE FTP MAYOR QUE FIP
	function ftpMayorFip(fip,ftp){
		var fip = fip.split("-");
		var ftp = ftp.split("-");
		//console.log(fip,ftp);
		var fip = new Date(fip[0],fip[1]-1,fip[2]);
		var ftp = new Date(ftp[0],ftp[1]-1,ftp[2])
		//console.log(fip,ftp);
		if (fip <= ftp) {
			return true;
		}else{
			return false;
		}
	}
	//VALIDO QUE LA FECHA SEA DEL CALENDARIO
	function existeFecha (fecha) {
        var fechaf = fecha.split("-");
        var d = fechaf[2];
        var m = fechaf[1];
        var y = fechaf[0];
        return m > 0 && m < 13 && y > 0 && y < 32768 && d > 0 && d <= (new Date(y, m, 0)).getDate();
	}
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
	<div class="data_table" style="/*position: relative;overflow: auto;width: 100%;*/">
		<table id= "lista-crud" class="table table-striped table-hover table-bordered table-condensed listar-act">
			<thead>
				<tr>
					<th>Fecha de Envío</th>
					<th>Enviado por</th>
					<th>Ubicación</th>
					<th>Equipo</th>
					<th>Identificación Equipo</th>
					<th>Sistema de bloqueo</th>
					<th><b>Trabajo</b></th>
					<th><b>SubTrabajo</b></th>
					<th>Fecha de Inicio Programada</th>
					<th>Fecha de Término Programada</th>
					<th>Fecha de Inicio Real</th>
					<th>Fecha de Término Real</th>
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
						$startTime->setTimezone(new DateTimeZone('America/Santiago'));
						echo $startTime->format('j F, Y, g:i a');
							?>
							<!--<div>
							<?php $uploaded= new DateTime($row['Entry']['CompleteTime']) ?>
								<span><b>Subido: </b>{{ $uploaded->format('d-F-Y g:i a') }}</span>
							</div>-->
						</td>
						<td><?php echo $row['Entry']['UserFirstName']." ".$row['Entry']['UserLastName'];?></td>
						<td><?php echo $row['EQUIPMENT']['LOCALIZATION_EQUIPMENT']['LOCALIZATION_NAME'];?></td>
						<td><?php echo $row['EQUIPMENT']['EQUIPMENT_NAME'];?></td>
						<td><?php echo $row['EQUIPMENT']['IDENTIFICATION_EQUIPMENT']['IDENTIFICATION_NAME'];?></td>
						<td><?php echo $row['EQUIPMENT']['BLOCK_SYSTEM'];?></td>
						<td><?php echo $row['EQUIPMENT']['WORK']['WORK_NAME'];?></td>
						<td><?php echo $row['EQUIPMENT']['WORK']['SUBWORK']['SUBWORK_NAME'];?></td>
						<td><?php echo $row['EQUIPMENT']['DATE_START_PROGRAMMED'];?></td>
						<td><?php echo $row['EQUIPMENT']['DATE_END_PROGRAMMED'];?></td>
						<td><?php
						$DATE_START_REAL = new DateTime($row['EQUIPMENT']['WORK']['SUBWORK']['DATE_START_REAL']);
						$DATE_START_REAL->setTimezone(new DateTimeZone('America/Santiago'));
						echo $DATE_START_REAL->format('j F, Y, g:i a');
						?></td>
						<td><?php
						$DATE_END_REAL = new DateTime($row['EQUIPMENT']['WORK']['SUBWORK']['DATE_END_REAL']);
						$DATE_END_REAL->setTimezone(new DateTimeZone('America/Santiago'));
						echo $DATE_END_REAL->format('j F, Y, g:i a');
						?></td>
						<td><?php echo $row['EQUIPMENT']['WORK']['SUBWORK']['POOP']."%";?></td>
						<td><?php echo $row['EQUIPMENT']['WORK']['SUBWORK']['OBSERVATIONS'];?></td>

						<td><?php

							if ($row['EQUIPMENT']['WORK']['PHOTOS']['PHOTO1']!= "") {
								$id = $row['Entry']['Id'];
								$Id = substr($id, 0, 8).'-'.substr($id, 8, 4).'-'.substr($id, 12, 4).'-'.substr($id, 16, 4).'-'.substr($id, 20, 32);
									echo '<a href="https://app.sendit.cl/Files/FormEntry/'.$row['ProviderId'].'-'.$Id.$row['EQUIPMENT']['WORK']['PHOTOS']['PHOTO1'].'">Ver Foto</a>';
							}else{echo "-";}
							?>


						</td>

					</tr>

				@endforeach

			</tbody>
		</table>
	</div>

@stop
</div><!--dataTableWrapper-->



