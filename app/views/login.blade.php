<!DOCTYPE html>
<html>
    <head>
        <meta charset="utf-8">
        <title>Login</title>
        {{ HTML::style('css/bootstrap.css'); }}
        <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.0.3/css/bootstrap.min.css" integrity="sha384-1q8mTJOASx8j1Au+a5WDVnPi2lkFfwwEAa8hDDdjZlpLegxhjVME1fgjWPGmkzs7" crossorigin="anonymous">
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
                            <label>
                                Recordar contraseña
                                {{ Form::checkbox('rememberme', true) }}
                            </label>
                        </div>
                        {{ Form::submit('Enviar', array('class' => 'btn btn-primary')) }}
                    {{ Form::close() }}
                </div>
            </div>
        </div>
        <script src="https://code.jquery.com/jquery.js"></script>
        {{ HTML::script('js/bootstrap.js'); }}
    </body>
</html>