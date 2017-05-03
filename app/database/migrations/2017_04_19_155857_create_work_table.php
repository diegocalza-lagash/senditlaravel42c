<?php

use Illuminate\Database\Schema\Blueprint;
use Illuminate\Database\Migrations\Migration;

class CreateWorkTable extends Migration {

	/**
	 * Run the migrations.
	 *
	 * @return void
	 */
	public function up()
	{
		Schema::create("works", function($tabla){
            $tabla->increments('id');
            $tabla->string('nombre', 100);
            $tabla->dateTime('fecha_inicio_programada');
            $tabla->dateTime('fecha_termino_programada');
            $tabla->string('sistema_bloqueo', 50);
            $tabla->string('s_turno_dia', 100);
            $tabla->string('s_turno_noche', 100);
            $tabla->string('i_p_turno_dia', 100);
            $tabla->string('i_p_turno_noche', 100);
            $tabla->string('foto1', 100);
            $tabla->string('descripcion_foto1', 100);
            $tabla->string('foto2', 100);
            $tabla->string('descripcion_foto2', 100);
            $tabla->string('video', 100);
            $tabla->string('descripcion_video', 100);

            $tabla->timestamps();
        });
	}

	/**
	 * Reverse the migrations.
	 *
	 * @return void
	 */
	public function down()
	{
		Schema::drop("works");
	}

}
