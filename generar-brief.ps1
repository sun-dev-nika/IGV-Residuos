$ErrorActionPreference = 'Stop'

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()

$doc.PageSetup.TopMargin = 36
$doc.PageSetup.BottomMargin = 36
$doc.PageSetup.LeftMargin = 50
$doc.PageSetup.RightMargin = 50

$selection = $word.Selection

$verdeOscuro = 2263842
$verdeMedio  = 3381811
$gris        = 6710886

function Add-Heading {
    param([string]$text, [int]$size = 14, [int]$color = 2263842, [bool]$bold = $true)
    $selection.Font.Name = 'Calibri'
    $selection.Font.Size = $size
    $selection.Font.Bold = $bold
    $selection.Font.Color = $color
    $selection.ParagraphFormat.SpaceBefore = 8
    $selection.ParagraphFormat.SpaceAfter = 4
    $selection.TypeText($text)
    $selection.TypeParagraph()
}

function Add-Text {
    param([string]$text, [int]$size = 10, [int]$color = 0, [bool]$bold = $false, [bool]$italic = $false)
    $selection.Font.Name = 'Calibri'
    $selection.Font.Size = $size
    $selection.Font.Bold = $bold
    $selection.Font.Italic = $italic
    $selection.Font.Color = $color
    $selection.ParagraphFormat.SpaceBefore = 0
    $selection.ParagraphFormat.SpaceAfter = 2
    $selection.TypeText($text)
    $selection.TypeParagraph()
}

function Add-Checklist {
    param([string]$label)
    $selection.Font.Name = 'Calibri'
    $selection.Font.Size = 10
    $selection.Font.Bold = $false
    $selection.Font.Color = 0
    $selection.ParagraphFormat.SpaceBefore = 0
    $selection.ParagraphFormat.SpaceAfter = 2
    $selection.ParagraphFormat.LeftIndent = 14
    $selection.TypeText([char]0x2610 + '  ' + $label)
    $selection.TypeParagraph()
    $selection.ParagraphFormat.LeftIndent = 0
}

function Add-Field {
    param([string]$label)
    $selection.Font.Name = 'Calibri'
    $selection.Font.Size = 10
    $selection.Font.Bold = $true
    $selection.Font.Color = 3381811
    if ($label.Length -gt 0) {
        $selection.TypeText($label + ': ')
    }
    $selection.Font.Bold = $false
    $selection.Font.Color = 0
    $selection.TypeText(('_' * 70))
    $selection.TypeParagraph()
}

# ENCABEZADO
$selection.ParagraphFormat.Alignment = 1
$selection.Font.Name = 'Calibri'
$selection.Font.Size = 26
$selection.Font.Bold = $true
$selection.Font.Color = $verdeOscuro
$selection.TypeText('IGV RESIDUOS')
$selection.TypeParagraph()

$selection.Font.Size = 12
$selection.Font.Bold = $false
$selection.Font.Italic = $true
$selection.Font.Color = $verdeMedio
$selection.TypeText([char]0x00AB + 'Brief para diseno de sitio web' + [char]0x00BB)
$selection.TypeParagraph()

$selection.Font.Italic = $false
$selection.Font.Size = 9
$selection.Font.Color = $gris
$selection.TypeText('Complete los campos y marque las casillas que correspondan. Puede dejar en blanco lo que aun no este definido.')
$selection.TypeParagraph()
$selection.TypeParagraph()
$selection.ParagraphFormat.Alignment = 0

# 1. INFORMACION CORPORATIVA
Add-Heading '1. INFORMACION CORPORATIVA'
Add-Field 'Razon social'
Add-Field 'Nombre comercial'
Add-Field 'RUT'
Add-Field 'Ano de fundacion'
Add-Field 'Sitio web actual'

Add-Text 'Mision:' 10 $verdeMedio $true
Add-Field ''
Add-Field ''

Add-Text 'Vision:' 10 $verdeMedio $true
Add-Field ''
Add-Field ''

Add-Text 'Valores corporativos:' 10 $verdeMedio $true
Add-Field ''
Add-Field ''

Add-Text 'Diferenciador frente a la competencia:' 10 $verdeMedio $true
Add-Field ''
Add-Field ''

Add-Text 'Certificaciones / autorizaciones - marque las que aplican:' 10 $verdeMedio $true
Add-Checklist 'ISO 14001'
Add-Checklist 'ISO 9001'
Add-Checklist 'Autorizacion Sanitaria SEREMI'
Add-Checklist 'Registro SINADER'
Add-Checklist 'Resolucion de Calificacion Ambiental / RCA'
Add-Checklist 'Registro RETC'
Add-Checklist 'Otras: ______________________________________________'

# 2. SERVICIOS
Add-Heading '2. SERVICIOS Y LINEAS DE NEGOCIO'

Add-Text 'Tipos de residuos que gestionan - marque:' 10 $verdeMedio $true
Add-Checklist 'Residuos industriales no peligrosos'
Add-Checklist 'Residuos peligrosos / RESPEL'
Add-Checklist 'Residuos solidos asimilables a domiciliarios'
Add-Checklist 'RAEE - residuos electronicos'
Add-Checklist 'Residuos de construccion y demolicion'
Add-Checklist 'Residuos organicos / compostables'
Add-Checklist 'Aceites y lubricantes usados'
Add-Checklist 'Chatarra y metales'
Add-Checklist 'Plasticos / carton / papel'
Add-Checklist 'Otros: ______________________________________________'

Add-Text 'Servicios que ofrecen - marque:' 10 $verdeMedio $true
Add-Checklist 'Retiro y transporte'
Add-Checklist 'Disposicion final'
Add-Checklist 'Valorizacion / reciclaje'
Add-Checklist 'Tratamiento de RESPEL'
Add-Checklist 'Arriendo de contenedores'
Add-Checklist 'Asesoria ambiental'
Add-Checklist 'Capacitaciones'
Add-Checklist 'Declaracion SINADER para clientes'
Add-Checklist 'Otros: ______________________________________________'

Add-Text 'Cobertura geografica - regiones y comunas:' 10 $verdeMedio $true
Add-Field ''
Add-Field ''

Add-Text 'Flota y equipamiento - camiones / tipos / plantas:' 10 $verdeMedio $true
Add-Field ''
Add-Field ''

# 3. IDENTIDAD VISUAL
Add-Heading '3. IDENTIDAD VISUAL'
Add-Checklist 'Logo en alta resolucion - PNG fondo transparente'
Add-Checklist 'Logo vectorial - AI / SVG / EPS'
Add-Checklist 'Manual de marca'
Add-Checklist 'Tipografias corporativas definidas'

Add-Field 'Color principal - HEX'
Add-Field 'Color secundario'
Add-Field 'Color de acento'
Add-Field 'Tipografia titulo'
Add-Field 'Tipografia cuerpo'

# 4. CONTENIDO VISUAL
Add-Heading '4. CONTENIDO VISUAL'
Add-Checklist 'Fotos reales de operaciones / camiones'
Add-Checklist 'Fotos de instalaciones / plantas'
Add-Checklist 'Fotos del equipo de trabajo con EPP'
Add-Checklist 'Videos corporativos o de proceso'
Add-Checklist 'Logos de clientes destacados - con autorizacion'
Add-Checklist 'Banco de imagenes propias en alta resolucion'

Add-Text 'Notas sobre material visual:' 10 $verdeMedio $true
Add-Field ''
Add-Field ''

# 5. PROYECTOS
Add-Heading '5. PROYECTOS Y CASOS DE EXITO'
Add-Text 'Liste 4 a 8 proyectos destacados. Por cada uno: cliente / descripcion / resultados.' 10 $gris $false $true

Add-Field 'Proyecto 1'
Add-Field ''
Add-Field 'Proyecto 2'
Add-Field ''
Add-Field 'Proyecto 3'
Add-Field ''
Add-Field 'Proyecto 4'
Add-Field ''

Add-Text 'Testimonios de clientes - nombre / cargo / empresa / frase:' 10 $verdeMedio $true
Add-Field ''
Add-Field ''

# 6. CONTACTO
Add-Heading '6. CONTACTO Y UBICACION'
Add-Field 'Direccion fisica'
Add-Field 'Telefono fijo'
Add-Field 'WhatsApp comercial'
Add-Field 'Correo de contacto general'
Add-Field 'Correo comercial / cotizaciones'
Add-Field 'Horarios de atencion'

Add-Text 'Redes sociales:' 10 $verdeMedio $true
Add-Field 'Instagram'
Add-Field 'LinkedIn'
Add-Field 'Facebook'
Add-Field 'YouTube / TikTok'

Add-Text 'Formulario de contacto - campos a capturar:' 10 $verdeMedio $true
Add-Checklist 'Nombre'
Add-Checklist 'Empresa / razon social'
Add-Checklist 'Correo'
Add-Checklist 'Telefono'
Add-Checklist 'Tipo de residuo a gestionar'
Add-Checklist 'Volumen estimado - kg / m3 / ton'
Add-Checklist 'Frecuencia - unica vez / mensual / etc'
Add-Checklist 'Comuna / ubicacion del retiro'
Add-Checklist 'Mensaje libre'

# 7. DOCUMENTOS
Add-Heading '7. DOCUMENTOS DESCARGABLES'
Add-Checklist 'Catalogo / brochure corporativo - PDF'
Add-Checklist 'Politica ambiental'
Add-Checklist 'Certificados ejemplo'
Add-Checklist 'Ficha tecnica de servicios'
Add-Checklist 'Otros: ______________________________________________'

# 8. TECNICOS
Add-Heading '8. ASPECTOS TECNICOS DEL SITIO'
Add-Field 'Dominio deseado - .cl'
Add-Checklist 'Dominio ya registrado'
Add-Checklist 'Necesita asesoria para registrarlo'
Add-Field 'Hosting actual'
Add-Field 'Correos corporativos necesarios'

Add-Text 'Integraciones requeridas:' 10 $verdeMedio $true
Add-Checklist 'Google Analytics'
Add-Checklist 'Google Tag Manager'
Add-Checklist 'Meta Pixel - Facebook / Instagram'
Add-Checklist 'WhatsApp Business - boton flotante'
Add-Checklist 'CRM - HubSpot / Salesforce / otro'
Add-Checklist 'Mailchimp / boletin'
Add-Checklist 'Google Maps embebido'

Add-Text 'Palabras clave para posicionamiento SEO:' 10 $verdeMedio $true
Add-Field ''
Add-Field ''

Add-Text 'Sitios web de referencia / inspiracion:' 10 $verdeMedio $true
Add-Field ''
Add-Field ''

Add-Text 'Sitios web de la competencia:' 10 $verdeMedio $true
Add-Field ''
Add-Field ''

# 9. SECCIONES
Add-Heading '9. SECCIONES DEL SITIO - marque las que desea incluir'
Add-Checklist 'Inicio / Hero con llamado a la accion'
Add-Checklist 'Nosotros / Quienes somos'
Add-Checklist 'Servicios'
Add-Checklist 'Tipos de residuos que gestionamos'
Add-Checklist 'Proceso de trabajo - paso a paso'
Add-Checklist 'Proyectos / casos de exito'
Add-Checklist 'Clientes / partners'
Add-Checklist 'Certificaciones y cumplimiento normativo'
Add-Checklist 'Sostenibilidad / impacto ambiental'
Add-Checklist 'Blog / noticias'
Add-Checklist 'Catalogo descargable'
Add-Checklist 'Cotizacion en linea'
Add-Checklist 'Preguntas frecuentes'
Add-Checklist 'Contacto con formulario y mapa'
Add-Checklist 'Trabaja con nosotros'

# 10. NOTAS
Add-Heading '10. NOTAS ADICIONALES Y COMENTARIOS LIBRES'
for ($i = 1; $i -le 8; $i++) {
    Add-Field ''
}

# PIE
$selection.TypeParagraph()
$selection.ParagraphFormat.Alignment = 1
$selection.Font.Size = 9
$selection.Font.Italic = $true
$selection.Font.Color = $gris
$selection.TypeText('Gracias por completar este brief. La informacion entregada nos permitira disenar un sitio web alineado a la identidad y objetivos de IGV Residuos.')
$selection.TypeParagraph()

$outPath = 'D:\joako\Paginas web\IGV RESIDUOS\Brief-IGV-Residuos.docx'
$doc.SaveAs([ref]$outPath, [ref]16)
$doc.Close()
$word.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($selection) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

Write-Output ('Archivo generado: ' + $outPath)
