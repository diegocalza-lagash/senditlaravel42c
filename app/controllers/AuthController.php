<?php

class AuthController extends BaseController {
    /*
    |--------------------------------------------------------------------------
    | Controlador de la autenticación de usuarios
    |--------------------------------------------------------------------------
    */
    /**
     * Muestra el formulario para login.
     */
    public function showLogin()
    {
        // Verificamos que el usuario no esté autenticado
        if (Auth::check())
        {
            // Si está autenticado lo mandamos a la raíz donde estara el mensaje de bienvenida.

            return Redirect::to('/');
        }else{
            // Mostramos la vista login.blade.php (Recordemos que .blade.php se omite.)
            return View::make('login');
        }


    }
    /**
     * Valida los datos del usuario.
     */
    public function postLogin()
    {
        // Guardamos en un arreglo los datos del usuario.
        $userdata = array(
            'username' => Input::get('username'),
            'password'=> Input::get('password')
        );
        // Validamos los datos y además mandamos como un segundo parámetro la opción de recordar el usuario.
        if(Auth::attempt($userdata))
        {
            // De ser datos válidos nos mandara a la bienvenida
            return Redirect::intended('dataform');
            //return View::make('DataSend.index'); //falla el make
           // return Redirect::route('home')
             //   ->with('flash_notice', 'You are successfully logged in.');
        }else{
            return Redirect::to('/login')
                    ->with('mensaje_error', 'Tus datos son incorrectos')
                    ->withInput();
        }
       //$username = Input::get('username');
        //$password = Input::get('password');
        //if (condition) {

        //}
        // En caso de que la autenticación haya fallado manda un mensaje al formulario de login y también regresamos los valores enviados con withInput().

    }
    /**
     * Muestra el formulario de login mostrando un mensaje de que cerró sesión.
     */
    public function logOut()
    {
        Auth::logout();
        Session::flush();
        return Redirect::to('/login')
                    ->with('mensaje_error', 'Tu sesión ha sido cerrada.');
    }
}