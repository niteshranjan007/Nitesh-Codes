<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.0/jquery.min.js"></script>
<script type="text/javascript">
$(document).ready(function(){
	$('.boldText').click(function(){
	   $('.container').toggleClass("bold");
	});
	$('.italicText').click(function(){
	  $('.container').toggleClass("italic");
	});
	$('.underlineText').click(function(){
	  $('.container').toggleClass("underline");
	});



	});</script>
<title>Insert title here</title>
<style type="text/css">
div.container {
    width: 300px;
    height: 100px;
    border: 1px solid #ccc;
    padding: 5px;
}
.bold{
  font-weight:bold;
}
.italic{
  font-style :italic;
}
.underline{
 text-decoration: underline;
}</style>
</head>
<body>
<div class="container" contentEditable></div><br/>
<input type="button" class="boldText" value="Bold">
<input type="button" class="italicText" value="Italic">
<input type="button" class="underlineText" value="Underline">
</body>
</html>