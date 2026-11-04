<?php
// Mantenemos la visualización de errores desactivada para producción
error_reporting(0);

$lifetime = 43200; // 12 horas en segundos
session_set_cookie_params($lifetime);
session_start();

// Validar que el usuario esté logueado
if (!isset($_SESSION['user_id'])) {
    session_destroy();
    header("location: login.php");
    exit;
}

// Conexión centralizada a la base de datos (UNA SOLA VEZ)
require_once 'conexion.php';
$mysqli = new mysqli($hostname, $username, $password, $database);

if ($mysqli->connect_error) {
    die("Error de conexión: " . $mysqli->connect_error);
}

// Variables de sesión consolidadas
$nivel = $_SESSION['nivel'];
$id_usuario = $_SESSION['user_id'];
$usuarioactivo = $_SESSION['firstname'] . " " . $_SESSION['lastname']; 
$coloricono = 0;

// Recibir y SANITIZAR la búsqueda (Previene Inyección SQL)
$busqueda_raw = $_REQUEST['busqueda'] ?? '';
$busqueda = $mysqli->real_escape_string(trim(strtoupper($busqueda_raw)));
$busqueda_html = htmlspecialchars($busqueda); // Para mostrar en el HTML sin romperlo

function eliminar_acentos($cadena){
    $cadena = str_replace(
        array('Á', 'À', 'Â', 'Ä', 'á', 'à', 'ä', 'â', 'ª', 'É', 'È', 'Ê', 'Ë', 'é', 'è', 'ë', 'ê', 'Í', 'Ì', 'Ï', 'Î', 'í', 'ì', 'ï', 'î', 'Ó', 'Ò', 'Ö', 'Ô', 'ó', 'ò', 'ö', 'ô', 'Ú', 'Ù', 'Û', 'Ü', 'ú', 'ù', 'ü', 'û'),
        array('A', 'A', 'A', 'A', 'a', 'a', 'a', 'a', 'a', 'E', 'E', 'E', 'E', 'e', 'e', 'e', 'e', 'I', 'I', 'I', 'I', 'i', 'i', 'i', 'i', 'O', 'O', 'O', 'O', 'o', 'o', 'o', 'o', 'U', 'U', 'U', 'U', 'u', 'u', 'u', 'u'),
        $cadena
    );
    return $cadena;
}

$cadena = strtr($busqueda, " ", "%");
$busquedav = eliminar_acentos($cadena);

date_default_timezone_set("America/Caracas");
$fechamod = date('d-m-Y') . " Hora: " . date('h:i A');

// Registrar la búsqueda en "chismoso"
$query100 = "INSERT INTO chismoso (evento, id_usuario, usuario, fecha) VALUES ('$busqueda', '$id_usuario', '$usuarioactivo', '$fechamod')";
$mysqli->query($query100);

// Obtener fondo de pantalla del usuario
$sql = "SELECT idfondo FROM users WHERE user_id='$id_usuario'";
$resultado_fondo = $mysqli->query($sql);
$idfondov = ($resultado_fondo && $resultado_fondo->num_rows > 0) ? $resultado_fondo->fetch_assoc()['idfondo'] : '1';

// Lógica de Búsqueda Unificada
$texto_limpio = preg_replace('/\b(DE|EL|LA|Y)\b/i', '', $busqueda);
$terminos = array_filter(explode(' ', $texto_limpio));

$condiciones_desc = '';
$condiciones_cod = '';

if (count($terminos) > 0) {
    foreach ($terminos as $termino) {
        $t = $mysqli->real_escape_string($termino);
        $condiciones_desc .= "descripcion LIKE '%$t%' AND ";
        $condiciones_cod .= "codigo LIKE '%$t%' OR ";
    }
    $condiciones_desc = rtrim($condiciones_desc, ' AND ');
    $condiciones_cod = rtrim($condiciones_cod, ' OR ');
} else {
    $condiciones_desc = "descripcion LIKE '%$busqueda%'";
    $condiciones_cod = "codigo LIKE '%$busqueda%'";
}

// Fuente nativa del sistema para optimizar rendimiento y legibilidad
$fuente = "system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif";
?>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no" />
    <title>Búsqueda - RSV</title>
    <link rel="icon" href="images/icono-color.png">
    <link rel="stylesheet" href="css/searchproductos<?php print htmlspecialchars($idfondov); ?>.style.css">
    <script src="js/all.min.js" crossorigin="anonymous"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.1.0/jquery.min.js"></script>
    
    <script>
        // Filtrado dinámico en la tabla
        $(document).ready(function(){
            $("#search").keyup(function(){
                _this = this;
                $.each($("#mytable tbody"), function() {
                    if($(this).text().toLowerCase().indexOf($(_this).val().toLowerCase()) === -1)
                        $(this).hide();
                    else
                        $(this).show();
                });
            });
        });
    </script>

    <style>
        /* Estilos del Modal Lightbox para Imágenes */
        .modal-img {
            display: none; 
            position: fixed;
            z-index: 9999;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.98);
            justify-content: center;
            align-items: center;
            flex-direction: column;
        }
        .modal-content {
            max-width: 90%;
            max-height: 80vh;
            border-radius: 8px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.5);
            object-fit: contain;
        }
        .close-modal {
            position: absolute;
            top: 20px;
            right: 30px;
            color: #ffffff;
            font-size: 40px;
            font-weight: bold;
            cursor: pointer;
        }
    </style>
</head>
<body style="font-family: <?php print $fuente; ?>;">
    <center>
        <table width="100%" border="0">
            <tr>
                <td align="left">
                    <a href="background.php?idusuario=<?php print urlencode($id_usuario); ?>&idfondo=<?php print urlencode($idfondov); ?>"><img src="svgs/solid/sun.svg" width="40" style="filter: invert(<?php print htmlspecialchars($idfondov); ?>)"></a>
                </td>
                <td align="right">
                    <a href="logout.php"><img src="svgs/solid/power-off.svg" width="40" style="filter: invert(<?php print htmlspecialchars($idfondov); ?>)"></a>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2">
                    <img id="miImagen" src="images/logo<?php print htmlspecialchars($idfondov); ?>.webp" width="500" onclick="window.location.href='dashboard.php'" style="cursor:pointer;">
                </td>
            </tr>
        </table>
    </center>

    <form method="post" action="searchproductos.php" name="registro" id="registro">
        <table width="200" align="center">
            <tr>
                <td>
                    <input type="text" class="btn-block" id="busqueda" name="busqueda" placeholder="Búsqueda" onkeyup="javascript:this.value=this.value.toUpperCase();" value="<?php print $busqueda_html; ?>" autofocus required>
                </td>
            </tr>
            <tr>
                <td>
                    <button type="submit" id="buscar">
                        <i class="fa-solid fa-car"></i>
                        <span>BUSCAR</span>
                    </button>
                </td>
            </tr>
        </table>
    </form>

    <p style="text-align:center;">Resultados para: <b><?php print $busqueda_html; ?></b></p>
    <br />

    <?php
    // ==========================================
    // 1. PRODUCTOS RSV / RSG
    // ==========================================
    $query = "SELECT codigo, montaje, name_m, existencia, existencia2, descripcion, imagen, preciog, nombre_almacen
              FROM ruedqlmd_rsvcomve_rsv.productos
              LEFT JOIN ruedqlmd_rsvcomve_rsv.marca ON productos.marca = marca.id_marca
              LEFT JOIN ruedqlmd_rsvcomve_rsv.almacenes ON productos.almacen = almacenes.id_almacen
              WHERE (($condiciones_desc) AND existencia > '0') 
                 OR (($condiciones_cod) AND existencia > '0') 
                 OR (($condiciones_desc) AND existencia2 > '0') 
                 OR (($condiciones_cod) AND existencia2 > '0') 
              ORDER BY productos.preciog DESC";
              
    $resultado = $mysqli->query($query);
    $rowcount = ($resultado) ? $resultado->num_rows : 0;
    ?>
    <center>
        <table width="30%" border='0'>
            <tr>
                <td align="center">
                    <?php if ($rowcount > 0): ?>
                        <font><?php print $rowcount; ?> Resultado(s) Principal(es)</font>
                    <?php else: ?>
                        <center><img src="images/advertencia.png" width="80"></center>
                        <font>Disculpa, en este momento no tenemos disponibles en nuestros almacenes.<br><br>Acá tienes el resultado de nuestros principales proveedores:</font>
                    <?php endif; ?>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <?php if ($rowcount > 0): ?>
                        <center><input type="text" id="search" placeholder="Filtrar resultados locales..."></center>
                    <?php endif; ?>
                    <br>
                </td>
            </tr>
        </table>
    </center>
    <br>
    
    <?php
    if ($rowcount > 0) {
        while ($rows = $resultado->fetch_assoc()) {
            $codigo = $rows['codigo'];
            $imagen = empty($rows['imagen']) ? "images/sinfoto.png" : $rows['imagen'];
            $existencia = $rows['existencia'];
            $existencia2 = $rows['existencia2'];
            $descripcion = $rows['descripcion'];
            $preciog = $rows['preciog'];
            $preciog2 = ($preciog * 1.16);
            
            $posicionRueda = '0';
            if (strpos($descripcion, "DELANTERO DERECHO") !== false) $posicionRueda = '2';
            elseif (strpos($descripcion, "DELANTERO IZQUIERDO") !== false) $posicionRueda = '1';
            elseif (strpos($descripcion, "TRASERO IZQUIERDO") !== false) $posicionRueda = '3';
            elseif (strpos($descripcion, "TRASERO DERECHO") !== false) $posicionRueda = '4';
            elseif (strpos($descripcion, "DELANTERO") !== false) $posicionRueda = '5';
            elseif (strpos($descripcion, "TRASERO") !== false) $posicionRueda = '6';
    ?>
    <table width="90%" border="0" align="center" id="mytable">
        <tr>
            <td colspan="3" class="titulo">
                <b><p><?php print htmlspecialchars($descripcion); ?></p></b> 
            </td>
            <td align="center">
                <?php if($nivel <= '1'): ?>
                    <a href="editardescripcion.php?producto=<?php print urlencode($codigo); ?>"><img src="img/edit.svg" width="40"></a>
                <?php endif; ?>
            </td>
        </tr>
        <tr>
            <td rowspan="5" width="10%" align="center">
                <img src="<?php print htmlspecialchars($imagen); ?>" 
                     style="max-width: 100px; width: 100%; height: auto; object-fit: contain; cursor: pointer;" 
                     loading="lazy" 
                     alt="<?php print htmlspecialchars($descripcion); ?>"
                     onclick="openModal('<?php print htmlspecialchars($imagen); ?>')">
            </td>
            <td colspan="2"></td>
            <td rowspan="5" width="10%" align="center"><img src="montaje/<?php print $posicionRueda; ?>.PNG" width="50"></td>
        </tr>
        <tr>
            <td class="peque" width="70%">Código: <font color="F76B2C" size="3"><?php print htmlspecialchars($codigo); ?></font></td>
        </tr>
        <tr>
            <td class="peque" width="70%">
                <font color="1BBCEE" size="3">GROUP:</font> <?php print $existencia; ?> Disponible(S)
            </td>
        </tr>
        <tr>
            <td class="peque" width="70%">
                <font color="194C82" size="3">VENEZUELA:</font> <?php print $existencia2; ?> Disponible(S)
            </td>
        </tr>
        <tr>
            <td class="peque" width="70%">Precio (Con IVA): <b><?php print number_format($preciog2, 2); ?>$</b> </td>
        </tr>
        <tr>
            <?php if (strpos($descripcion, 'PROLED') !== false): ?>
                <td align="center"><a href="pdf/catalogoproled.pdf"><p style="color:green;">Catálogo ProLED</p></a></td>
                <td align="center"><a href="pdf/aplicacionproled.pdf"><p style="color:green;">Aplicaciones ProLED</p></a></td>
            <?php endif; ?>
        </tr>
        <tr>
            <td colspan="4"><hr style="width:95%;height:2px;color:#00FF01;background-color:#00FF01"></td>
        </tr>
    </table>
    <?php
        }
    }
    ?>

    <!-- ================== CARVICA ================== -->
    <?php
    $query3 = "SELECT codigo, descripcion, marca, existencia, precio
               FROM ruedqlmd_rsvcomve_rsv.carvica
               WHERE ($condiciones_desc) OR ($condiciones_cod) OR (marca LIKE '%$busqueda%') 
               ORDER BY carvica.precio DESC";
    $resultado3 = $mysqli->query($query3);
    $rowcount3 = ($resultado3) ? $resultado3->num_rows : 0;
    
    if ($rowcount3 > 0):
    ?>
    <br>
    <center>
        <b><h1>LISTA DE PRECIOS CARVICA</h1></b>
        <h3 style="color:red;">Existencia referencial (Consultar disponibilidad)</h3>
        <font><?php print $rowcount3; ?> Resultado(s)</font>
    </center>
    <?php
        while ($rows3 = $resultado3->fetch_assoc()) {
            $precio3coniva = ($rows3['precio'] * 1.25) * 1.16;
    ?>
    <table width="90%" border="0" align="center">
        <tr>
            <td colspan="3" class="titulo">
                <b><p><?php print htmlspecialchars($rows3['descripcion']); ?></p></b> 
            </td>
        </tr>
        <tr><td class="peque" width="70%">Código: <font size="3"><?php print htmlspecialchars($rows3['codigo']); ?></font></td></tr>
        <tr><td class="peque" width="70%">Marca: <font size="3"><?php print htmlspecialchars($rows3['marca']); ?></font></td></tr>
        <tr><td class="peque" width="70%">Existencia: <font size="3"><?php print htmlspecialchars($rows3['existencia']); ?></font></td></tr>
        <tr><td class="peque" width="70%">Precio (Con IVA): <b><?php print number_format($precio3coniva, 2); ?>$</b></td></tr>
        <tr><td colspan="4"><hr style="width:95%;height:2px;color:#E20012;background-color:#E20012"></td></tr>
    </table>
    <?php } endif; ?>

    <!-- ================== MHW ================== -->
    <?php
    $query6 = "SELECT codigo, descripcion, marca, precio
               FROM ruedqlmd_rsvcomve_rsv.mhw
               WHERE ($condiciones_desc) OR ($condiciones_cod) OR (marca LIKE '%$busqueda%') 
               ORDER BY mhw.precio DESC";
    $resultado6 = $mysqli->query($query6);
    $rowcount6 = ($resultado6) ? $resultado6->num_rows : 0;
    
    if ($rowcount6 > 0):
    ?>
    <br>
    <center>
        <b><h1>LISTA DE PRECIOS MHW</h1></b>
        <h3 style="color:red;">Precio Referencial (Consultar disponibilidad)</h3>
        <font><?php print $rowcount6; ?> Resultado(s)</font>
    </center>
    <?php
        while ($rows6 = $resultado6->fetch_assoc()) {
            $precio6coniva = ($rows6['precio'] * 1.25) * 1.16;
    ?>
    <table width="90%" border="0" align="center">
        <tr>
            <td colspan="3" class="titulo">
                <b><p><?php print htmlspecialchars($rows6['descripcion']); ?></p></b> 
            </td>
        </tr>
        <tr><td class="peque" width="70%">Código: <font size="3"><?php print htmlspecialchars($rows6['codigo']); ?></font></td></tr>
        <tr><td class="peque" width="70%">Marca: <font size="3"><?php print htmlspecialchars($rows6['marca']); ?></font></td></tr>
        <tr><td class="peque" width="70%">Precio (Con IVA): <b><?php print number_format($precio6coniva, 2); ?>$</b></td></tr>
        <tr><td colspan="4"><hr style="width:95%;height:2px;color:#FF7000;background-color:#FF7000"></td></tr>
    </table>
    <?php } endif; ?>

    <!-- ================== AUTOPARTES 2006 ================== -->
    <?php
    $query4 = "SELECT codigo, descripcion, marca, existencia, precio
               FROM ruedqlmd_rsvcomve_rsv.autopartes2006
               WHERE ($condiciones_desc) OR ($condiciones_cod) OR (marca LIKE '%$busqueda%') 
               ORDER BY autopartes2006.precio DESC";
    $resultado4 = $mysqli->query($query4);
    $rowcount4 = ($resultado4) ? $resultado4->num_rows : 0;
    
    if ($rowcount4 > 0):
    ?>
    <br>
    <center>
        <b><h1>LISTA DE PRECIOS AUTOPARTES 2006</h1></b>
        <h3 style="color:red;">Existencia referencial (Consultar disponibilidad)</h3>
        <font><?php print $rowcount4; ?> Resultado(s)</font>
    </center>
    <?php
        while ($rows4 = $resultado4->fetch_assoc()) {
            $precio4coniva = ($rows4['precio'] * 1.25) * 1.16;
    ?>
    <table width="90%" border="0" align="center">
        <tr>
            <td colspan="3" class="titulo">
                <b><p><?php print htmlspecialchars($rows4['descripcion']); ?></p></b> 
            </td>
        </tr>
        <tr><td class="peque" width="70%">Código: <font size="3"><?php print htmlspecialchars($rows4['codigo']); ?></font></td></tr>
        <tr><td class="peque" width="70%">Marca: <font size="3"><?php print htmlspecialchars($rows4['marca']); ?></font></td></tr>
        <tr><td class="peque" width="70%">Existencia: <font size="3"><?php print htmlspecialchars($rows4['existencia']); ?></font></td></tr>
        <tr><td class="peque" width="70%">Precio (Con IVA): <b><?php print number_format($precio4coniva, 2); ?>$</b></td></tr>
        <tr><td colspan="4"><hr style="width:95%;height:2px;color:#005FB8;background-color:#005FB8"></td></tr>
    </table>
    <?php } endif; ?>

    <!-- ================== LECOMP ================== -->
    <?php
    $query2 = "SELECT codigo, descripcion, marca, precio
               FROM ruedqlmd_rsvcomve_rsv.lecomp
               WHERE ($condiciones_desc) OR ($condiciones_cod) OR (marca LIKE '%$busqueda%') 
               ORDER BY lecomp.precio DESC";
    $resultado2 = $mysqli->query($query2);
    $rowcount2 = ($resultado2) ? $resultado2->num_rows : 0;
    
    if ($rowcount2 > 0):
    ?>
    <br>
    <center>
        <b><h1>LISTA DE PRECIOS LECOMP</h1></b>
        <h3 style="color:red;">Precios referenciales (Consultar disponibilidad)</h3>
        <font><?php print $rowcount2; ?> Resultado(s)</font>
    </center>
    <?php
        while ($rows2 = $resultado2->fetch_assoc()) {
            $precio2coniva = ($rows2['precio'] * 1.25) * 1.16;
    ?>
    <table width="90%" border="0" align="center">
        <tr>
            <td colspan="3" class="titulo">
                <b><p><?php print htmlspecialchars($rows2['descripcion']); ?></p></b> 
            </td>
        </tr>
        <tr><td class="peque" width="70%">Código: <font size="3"><?php print htmlspecialchars($rows2['codigo']); ?></font></td></tr>
        <tr><td class="peque" width="70%">Marca: <font size="3"><?php print htmlspecialchars($rows2['marca']); ?></font></td></tr>
        <tr><td class="peque" width="70%">Precio (Con IVA): <b><?php print number_format($precio2coniva, 2); ?>$</b></td></tr>
        <tr><td colspan="4"><hr style="width:95%;height:2px;color:#23AAF2;background-color:#23AAF2"></td></tr>
    </table>
    <?php } endif; ?>

    <br><br><br><br>

    <!-- Estructura del Modal -->
    <div id="imageModal" class="modal-img" onclick="closeModal()">
        <span class="close-modal">&times;</span>
        <img class="modal-content" id="imgModalCentro">
    </div>

    <!-- Script para abrir y cerrar el modal -->
    <script>
        function openModal(imgSrc) {
            document.getElementById("imageModal").style.display = "flex";
            document.getElementById("imgModalCentro").src = imgSrc;
        }
        function closeModal() {
            document.getElementById("imageModal").style.display = "none";
            document.getElementById("imgModalCentro").src = "";
        }
    </script>
</body>
</html>