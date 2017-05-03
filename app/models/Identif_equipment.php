<?php

	class Identif_equipment extends Eloquent
	{
		protected $table = 'identif_equipments';


    	 public function equipment()
        {
            return $this->belongsTo('Equipment');
        }
         public function localizations()
        {
            return $this->hasMany('Localization');
        }

    }
?>