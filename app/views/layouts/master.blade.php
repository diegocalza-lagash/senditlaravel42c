<!DOCTYPE html>
<html ng-app>
<head>
	<meta charset="utf-8">
	<title>@yield('title')</title>

	  <!--<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/0.98.2/css/materialize.min.css">

	  <script src="https://cdnjs.cloudflare.com/ajax/libs/materialize/0.98.2/js/materialize.min.js"></script>-->
	  <!--<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">
	  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap-theme.min.css" integrity="sha384-rHyoN1iRsVXV4nD0JutlnGaslCJuC7uwjduW9SVrLvRYooPp2bWYgmgJQIXwl/Sp" crossorigin="anonymous">
		<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js" integrity="sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa" crossorigin="anonymous"></script>-->
	<link rel="stylesheet" href="{{ URL::asset('assets/css/bootstrap.min.css') }}">
	<link rel="stylesheet" href="{{ URL::asset('assets/css/bootstrap.css') }}">
	<link rel="stylesheet" href="{{ URL::asset('assets/css/css-table.css') }}">
	<link rel="stylesheet" href="{{ URL::asset('assets/css/reset.css') }}">
	<script src="http://code.jquery.com/jquery-latest.js"></script>
	<!--{{ HTML::script('assets/js/bootstrap.min.js') }}-->
	<script type="text/javascript" src="{{ URL::asset('assets/jquery-ui-1.12.1/jquery-ui.min.js') }}"></script>
	<script type="text/javascript" src="{{ URL::asset('assets/jquery-ui-1.12.1/jquery-ui.js') }}"></script>
	<script type="text/javascript" src="{{ URL::asset('assets/js/bootstrap.min.js') }}"></script>
	<script type="text/javascript" src="{{ URL::asset('assets/js/bootstrap.js') }}"></script>
	<script type="text/javascript" src="{{ URL::asset('assets/js/angular.min.js') }}"></script>


</head>
<body>
@section('nav')
	<header>
		<nav class="navbar navbar-inverse navbar-fixed-top">
			<div class="container-fluid">
				<div class="navbar-header">
					<button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#navbar" aria-expanded="false" aria-controls="navbar">
			            <span class="sr-only">Menu</span>
			            <span class="icon-bar"></span>
			            <span class="icon-bar"></span>
			            <span class="icon-bar"></span>
	          		</button>
				</div>
				<div id="navbar" class="navbar-collapse collapse">
		          <ul class="nav navbar-nav ">
		            <li><a href="#">Rudel Report</a></li>

		          </ul>
		          <form class="navbar-form navbar-right">
		            <input class="form-control" placeholder="Search..." type="text">
		          </form>
		        </div>
			</div>
		</nav>
	</header>
@show

    @section('sidebar')
    <div class= "col-sm-3 col-md-2 sidebar " style="padding-top: 2%;">
          <ul class="nav nav-sidebar">
          	@yield('sidebar')
          </ul>
          <ul class="nav nav-sidebar">
            @yield('trabajos')
          </ul>
          <ul class="nav nav-sidebar">
            <li><a href="/logout">Log Out</a></li>
          </ul>
           <ul class="nav nav-sidebar">

          </ul>

    </div>
	@show
	<div class="container-fluid">
		@yield('content')
	</div>

</body>
</html>
