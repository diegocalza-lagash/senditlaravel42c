<?php

	class Localization extends Eloquent
	{
		protected $table = 'localizations';


    	 public function identification()
        {
            return $this->belongsTo('Identif_equipment');
        }
        public function works()
        {
            return $this->hasMany('Work');
        }
    }
?>