<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<title>你好</title>
</head>
<body>
<?php 
// echo $_POST["x"];

function sum($x,$y) {
  $z=$x*$y;
  return $z;
}



echo $_POST["x"]."*".$_POST["y"]."的运算结果为:".sum($_POST["x"],$_POST["y"]);
?>
</body>
</html>