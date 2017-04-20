<?php

    class Subwork extends Eloquent
    {
        protected $table = 'subworks';


	     public function work()
	    {
	        return $this->belongsTo('Work');
	    }

	}
?>