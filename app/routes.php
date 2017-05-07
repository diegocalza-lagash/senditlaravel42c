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
 /*Route::get('/', function()
    {
        return View::make('login');
    });*/

Route::resource('data','DataSendController');
Route::controller('dataform','DataSendController');

Route::get('report/show', 'DataSendController@report');

//Route::get('list-works', 'DataSendController@showWorks');
Route::resource('excel','ExcelController');
// Nos mostrará el formulario de login.

//Route::get('/', array('as' => 'home', function () { }));
//Route::get('login', array('as' => 'login', function () { }))->before('guest');

Route::get('login', 'AuthController@showLogin');
// Validamos los datos de inicio de sesión.
Route::post('login', 'AuthController@postLogin');
//Route::post('login', ['uses' => 'AuthController@postLogin', 'before' => 'guest']);
Route::get('logout', 'AuthController@logOut');
//Route::get('/logout', ['uses' => 'AuthController@logOut', 'before' => 'auth']);


Route::get('/dataform', array('before' => 'auth', function(){
    return View::make('DataSend.index');
}));
Route::get('/', array('as' => 'home', function(){
    return View::make('DataSend.index');
}))->before('auth');
//for download excel
Route::get('/download','HomeController@getDownload');
//AJAX
Route::get('/getEquipments', 'DataSendController@getEquipments');