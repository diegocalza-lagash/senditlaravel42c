<?php

	class Equipment extends Eloquent
	{
		protected $table = 'equipments';

		public function identifications()
		{
			return $this -> hasMany('Identif_equipment');
		}
	}
?>