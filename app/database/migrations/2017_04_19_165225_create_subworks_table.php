<?php

use Illuminate\Database\Schema\Blueprint;
use Illuminate\Database\Migrations\Migration;

class CreateSubworksTable extends Migration {

	/**
	 * Run the migrations.
	 *
	 * @return void
	 */
	public function up()
	{
		Schema::create("subworks", function($tabla){
            $tabla->increments('id');
            $tabla->string('nombre', 100);
            $tabla->dateTime('fecha_inicio_real');
            $tabla->dateTime('fecha_termino_real');
            $tabla->integer('poop');
            $tabla->longText('observaciones');


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
		//
	}

}
