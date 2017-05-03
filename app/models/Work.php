<?php

	class Work extends Eloquent
	{
		protected $table = 'works';


    	 public function localization()
        {
            return $this->belongsTo('Localization');
        }
        public function subworks()
        {
            return $this->hasMany('Subwork');
        }
    }
?>