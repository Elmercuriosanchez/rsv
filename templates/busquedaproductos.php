<?php
session_start();

if (!isset($_SESSION['user_login_status']) || $_SESSION['user_login_status'] != 1) {
    header("location: login.php");
    exit;
}

require_once 'conexion.php';
$mysqli = new mysqli($hostname, $username, $password, $database);

if ($mysqli->connect_error) {
    die("Error de conexión: " . $mysqli->connect_error);
}

$id_usuario = $_SESSION['user_id'];
$nivel = $_SESSION['nivel'] ?? '';

// Método clásico soportado por tu servidor
$sql = "SELECT idfondo FROM users WHERE user_id = '$id_usuario'";
$resultado = $mysqli->query($sql);
$idfondov = '1';
if ($resultado && $resultado->num_rows > 0) {
    $row = $resultado->fetch_assoc();
    $idfondov = $row['idfondo'];
}

$query = "SELECT user_id, firstname, lastname FROM users WHERE user_id = '$id_usuario'";
$resultado_user = $mysqli->query($query);
$id = '';
$usuarioactivo = '';
if ($resultado_user && $resultado_user->num_rows > 0) {
    $rows = $resultado_user->fetch_assoc();
    $id = $rows['user_id'];
    $nombre = $rows['firstname'];
    $apellido = $rows['lastname'];
    $usuarioactivo = $nombre . " " . $apellido;
}
?>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no" />
    <title>Búsqueda Productos - RSV</title>
    <link rel="stylesheet" href="css/busquedaproductos<?php print htmlspecialchars($idfondov); ?>.style.css">
    <link rel="icon" href="images/icono-color.png">
    <script src="js/all.min.js" crossorigin="anonymous"></script>
    <script>
        function showHint(str) {
            if (str.length == 0) {
                document.getElementById("txtHint").innerHTML = "";
                return;
            } else {
                var xmlhttp = new XMLHttpRequest();
                xmlhttp.onreadystatechange = function() {
                    if (this.readyState == 4 && this.status == 200) {
                        document.getElementById("txtHint").innerHTML = this.responseText;
                    }
                };
                xmlhttp.open("GET", "gethint.php?q=" + str, true);
                xmlhttp.send();
            }
        }
    </script>
</head>
<body>

<div class="imagen-superior">
    <img id="miImagen" src="images/logo<?php print htmlspecialchars($idfondov); ?>.webp" width="500" onclick="window.location.href='dashboard.php'" style="cursor: pointer;">
</div>

<center>
    <form method="post" action="searchproductos.php?id=<?php print urlencode($id); ?>&usuario=<?php print urlencode($usuarioactivo); ?>" name="registro" id="registro">
        <table align="center">
            <tr>
                <td align="center">
                    <div class="search">
                        <input type="text" class="search__input" id="busqueda" name="busqueda" placeholder="Búsqueda" onkeyup="javascript:this.value=this.value.toUpperCase();" autofocus required>
                    </div>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <br>
                    <button type="submit" id="submit" name="login">Buscar</button>
                </td>
            </tr>
        </table>
    </form>
</center>

</body>
</html>