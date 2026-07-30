$ErrorActionPreference = 'Stop'

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()

$doc.PageSetup.TopMargin = 50
$doc.PageSetup.BottomMargin = 50
$doc.PageSetup.LeftMargin = 60
$doc.PageSetup.RightMargin = 60

$sel = $word.Selection

$verde       = 5135390   # ~ #1E8A4E reverse BGR
$verdeOscuro = 3300623   # ~ #0F5A33
$navy        = 4337434   # ~ #1A2942
$gris        = 7763574
$grisClaro   = 14342874  # ~ #DBDBDA

function H1 { param([string]$t)
    $sel.Font.Name = 'Calibri'; $sel.Font.Size = 22; $sel.Font.Bold = $true; $sel.Font.Color = $verdeOscuro
    $sel.ParagraphFormat.SpaceBefore = 0; $sel.ParagraphFormat.SpaceAfter = 6
    $sel.ParagraphFormat.Alignment = 1
    $sel.TypeText($t); $sel.TypeParagraph()
    $sel.ParagraphFormat.Alignment = 0
}
function H2 { param([string]$t)
    $sel.Font.Name = 'Calibri'; $sel.Font.Size = 14; $sel.Font.Bold = $true; $sel.Font.Color = $verde
    $sel.ParagraphFormat.SpaceBefore = 14; $sel.ParagraphFormat.SpaceAfter = 4
    $sel.TypeText($t); $sel.TypeParagraph()
}
function P { param([string]$t)
    $sel.Font.Name = 'Calibri'; $sel.Font.Size = 11; $sel.Font.Bold = $false; $sel.Font.Color = 0
    $sel.ParagraphFormat.SpaceBefore = 0; $sel.ParagraphFormat.SpaceAfter = 4
    $sel.TypeText($t); $sel.TypeParagraph()
}
function Step { param([int]$n, [string]$titulo, [string]$cuerpo)
    $sel.Font.Name = 'Calibri'; $sel.Font.Size = 12; $sel.Font.Bold = $true; $sel.Font.Color = $navy
    $sel.ParagraphFormat.SpaceBefore = 10; $sel.ParagraphFormat.SpaceAfter = 2
    $sel.TypeText(('Paso ' + $n + '. ' + $titulo))
    $sel.TypeParagraph()
    $sel.Font.Size = 11; $sel.Font.Bold = $false; $sel.Font.Color = 0
    $sel.ParagraphFormat.SpaceBefore = 0; $sel.ParagraphFormat.SpaceAfter = 6
    $sel.ParagraphFormat.LeftIndent = 14
    $sel.TypeText($cuerpo)
    $sel.TypeParagraph()
    $sel.ParagraphFormat.LeftIndent = 0
}
function Tip { param([string]$t)
    $sel.Font.Name = 'Calibri'; $sel.Font.Size = 10; $sel.Font.Italic = $true; $sel.Font.Color = $gris
    $sel.ParagraphFormat.SpaceBefore = 2; $sel.ParagraphFormat.SpaceAfter = 6
    $sel.ParagraphFormat.LeftIndent = 14
    $sel.TypeText([char]0x2192 + ' ' + $t)
    $sel.TypeParagraph()
    $sel.Font.Italic = $false
    $sel.ParagraphFormat.LeftIndent = 0
}

# PORTADA
$sel.ParagraphFormat.Alignment = 1
$sel.Font.Size = 10; $sel.Font.Color = $verde; $sel.Font.Bold = $true
$sel.TypeText('IGV RESIDUOS'); $sel.TypeParagraph()
$sel.Font.Size = 26; $sel.Font.Color = $navy
$sel.TypeText('Como actualizar la web')
$sel.TypeParagraph()
$sel.Font.Size = 12; $sel.Font.Bold = $false; $sel.Font.Italic = $true; $sel.Font.Color = $gris
$sel.TypeText('Guia paso a paso para agregar y editar Puntos Limpios desde Google Sheets')
$sel.TypeParagraph(); $sel.TypeParagraph()
$sel.ParagraphFormat.Alignment = 0
$sel.Font.Italic = $false

# INTRO
H2 'Como funciona'
P 'Tu pagina web lee la informacion de Puntos Limpios desde una planilla de Google Sheets. Cuando agregas, editas o eliminas una fila en la planilla, la web se actualiza sola al recargar - no necesitas saber nada de programacion.'

H2 'Lo que recibiste'
P ([char]0x2022 + ' Un link a tu planilla de Google Sheets (la abriste con tu cuenta Google).')
P ([char]0x2022 + ' Un link a una carpeta de Google Drive donde subir las fotos.')
P ([char]0x2022 + ' Esta guia.')

# PASO A PASO
H2 'Agregar un nuevo Punto Limpio'

Step 1 'Sube la foto del punto a Google Drive' 'Abre la carpeta de Drive compartida. Arrastra la foto (formato JPG o PNG, ideal 1200x800 px). Espera unos segundos a que termine de subir.'
Tip 'Si la foto sale del celular, mandatela por mail o WhatsApp a tu computador antes para evitar fotos muy pesadas.'

Step 2 'Copia el link de la foto' 'En Google Drive: click derecho sobre la foto > Compartir > Copiar link. IMPORTANTE: asegurate de que diga "Cualquiera con el link puede ver", si no, la foto no se vera en la web.'

Step 3 'Abre la planilla de Google Sheets' 'Hace click en el link de la planilla. Te aseguras de estar en la pestaña "puntos_limpios" (abajo de la planilla).'

Step 4 'Agrega una fila nueva al final' 'Baja hasta la primera fila vacia y completa de izquierda a derecha:'
P '   - nombre: nombre del lugar (ej: Stripcenter Las Brujas)'
P '   - comuna: comuna donde esta el punto (ej: La Reina)'
P '   - estado: escribir exactamente "Inaugurado", "Proximamente" o "En operacion"'
P '   - fecha_inauguracion: en formato AAAA-MM-DD (ej: 2026-03-15). Si todavia no hay dia, usa AAAA-MM (ej: 2026-07)'
P '   - descripcion: texto libre que describa el punto, una o dos frases'
P '   - residuos_aceptados: lista separada por comas, en minuscula (ej: plastico, vidrio, carton, metal, tetra, papel)'
P '   - foto_url: pega aqui el link que copiaste en el Paso 2'
P '   - activo: escribir "si" para que aparezca en la web'

Step 5 'Listo' 'Cierra la pestaña. Espera unos 5 segundos. Abre tu pagina web en el navegador y recarga (Ctrl+R o F5 en Windows). El nuevo punto limpio aparecera automaticamente.'

H2 'Editar un Punto Limpio existente'
P 'Solo cambia la celda que necesitas modificar (ej: actualizar la descripcion, cambiar el estado de "Proximamente" a "Inaugurado", actualizar la foto). Guarda y recarga la web.'

H2 'Ocultar un Punto Limpio sin borrarlo'
P 'Cambia la columna "activo" de "si" a "no". El punto deja de aparecer en la web pero la fila queda guardada por si quieres volver a mostrarlo despues.'

H2 'Eliminar un Punto Limpio'
P 'Selecciona la fila completa, click derecho > Eliminar fila. La web se actualiza al recargar.'

H2 'Cosas importantes'
P ([char]0x2022 + ' NO cambies los nombres de las columnas (primera fila). Si los cambias, la web deja de leer bien.')
P ([char]0x2022 + ' Los estados validos son SOLO: Inaugurado / Proximamente / En operacion.')
P ([char]0x2022 + ' Cada cambio tarda unos 5 segundos en propagarse de Google a la web.')
P ([char]0x2022 + ' Si la foto no aparece, vuelve al paso 2 y verifica los permisos de la foto en Drive.')
P ([char]0x2022 + ' Pueden editar la planilla varias personas a la vez (compartila desde el menu Compartir de Sheets).')

H2 'Problemas frecuentes'
P 'La foto no se ve en la web > revisa que el link sea de Google Drive y que tenga permiso "Cualquiera con el link puede ver".'
P 'El punto no aparece en la web > revisa que la columna activo diga "si" (sin espacios) y que la fila no este vacia.'
P 'La web tarda en mostrar el cambio > Google demora hasta 5 minutos en propagar cambios grandes. Recarga con Ctrl+Shift+R para limpiar cache.'

# PIE
$sel.TypeParagraph()
$sel.ParagraphFormat.Alignment = 1
$sel.Font.Size = 9; $sel.Font.Italic = $true; $sel.Font.Color = $gris
$sel.TypeText('Ante cualquier duda, escribi a tu desarrollador web.')
$sel.TypeParagraph()
$sel.Font.Italic = $false

$outPath = 'D:\joako\Paginas web\IGV RESIDUOS\docs\instructivo-cliente.docx'
$doc.SaveAs([ref]$outPath, [ref]16)
$doc.Close()
$word.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sel) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

Write-Output ('Instructivo generado: ' + $outPath)
