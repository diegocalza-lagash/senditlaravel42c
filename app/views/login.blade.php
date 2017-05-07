<!DOCTYPE html>
<html>
    <head>
        <meta charset="utf-8">
        <title>Login</title>

        <link rel="stylesheet" href="{{ URL::asset('assets/css/bootstrap.min.css') }}">
        <link rel="stylesheet" href="{{ URL::asset('assets/css/bootstrap.css') }}">
        <link rel="stylesheet" href="{{ URL::asset('assets/css/css-table.css') }}">
        <link rel="stylesheet" href="{{ URL::asset('assets/css/reset.css') }}">
        <script src="http://code.jquery.com/jquery-latest.js"></script>
        <!--{{ HTML::script('assets/js/bootstrap.min.js') }}-->
        <script type="text/javascript" src="{{ URL::asset('assets/js/bootstrap.min.js') }}"></script>
        <script type="text/javascript" src="{{ URL::asset('assets/js/bootstrap.js') }}"></script>
    </head>
    <body>
        <div class="container">
            <div class="panel panel-default">
                <div class="panel-body">
                    {{-- Preguntamos si hay algún mensaje de error y si hay lo mostramos  --}}
                    @if(Session::has('mensaje_error'))
                        <div class="alert alert-danger">{{ Session::get('mensaje_error') }}</div>
                    @endif
                    {{ Form::open(array('url' => '/login')) }}
                        <legend>Iniciar sesión</legend>
                        <div class="form-group">
                            {{ Form::label('usuario', 'Nombre de usuario') }}
                            {{ Form::text('username', Input::old('username'), array('class' => 'form-control')); }}
                        </div>
                        <div class="form-group">
                            {{ Form::label('contraseña', 'Contraseña') }}
                            {{ Form::password('password', array('class' => 'form-control')); }}
                        </div>
                        <div class="checkbox">
                        <p>Recordar contraseña</p>
                            <label>

                                {{ Form::checkbox('rememberme', true) }}
                            </label>
                        </div>
                        {{ Form::submit('Enviar', array('class' => 'btn btn-primary')) }}
                    {{ Form::close() }}
                </div>
            </div>
        </div>

    </body>
</html>