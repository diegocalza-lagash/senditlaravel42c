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
        return View::make('login');
    });
Route::resource('data','DataSendController');
Route::controller('dataform','DataSendController');

Route::get('report/show', 'DataSendController@report');
//Route::get('list-works', 'DataSendController@showWorks');
Route::resource('excel','ExcelController');
// Nos mostrará el formulario de login.
Route::get('login', 'AuthController@showLogin');
// Validamos los datos de inicio de sesión.
Route::post('login', 'AuthController@postLogin');
Route::get('logout', 'AuthController@logOut');

/*Route::get('dataform', array('before' => 'auth', function(){
    return View::make('DataSend.index');
}));*/


