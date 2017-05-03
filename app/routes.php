<?php
error_reporting(E_ALL);
/*
|--------------------------------------------------------------------------
| Application Routes
|--------------------------------------------------------------------------
|
| Here is where you can register all of the routes for an application.
| It's a breeze. Simply tell Laravel the URIs it should respond to
| and give it the Closure to execute when that URI is requested.
|
*/

Route::get('/', function()
{
	return View::make('hello');
});
Route::resource('data','DataSendController');
Route::controller('dataform','DataSendController');
//Route::get('dataform', array('uses' => 'DataSendController@showWorks'));
/*Route::get('dataform', function()
{

	return View::make('dataSends.index');
});*/
Route::get('report/show', 'DataSendController@report');
//Route::get('list-works', 'DataSendController@showWorks');
Route::resource('excel','ExcelController');