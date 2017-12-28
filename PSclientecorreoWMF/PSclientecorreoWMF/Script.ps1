#
# Script.ps1
#
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$xaml = @'
<Window x:Name="ClienteDeCorreo" 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        Title="Cliente de Correo PS 2017" Height="370.588" Width="518.824">
    <Grid x:Name="wdwMain" Margin="0,0,18,12">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="63*"/>
            <ColumnDefinition Width="430*"/>
        </Grid.ColumnDefinitions>
        <Label x:Name="lblFrom" Content="De:" HorizontalAlignment="Left" Height="27" Margin="23,23,0,0" VerticalAlignment="Top" Width="49" Grid.ColumnSpan="2"/>
        <TextBox x:Name="txtFrom" HorizontalAlignment="Left" Height="27" Margin="60,23,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="423" Grid.ColumnSpan="2"/>
        <Label x:Name="lblTo" Content="Para:" HorizontalAlignment="Left" Height="27" Margin="23,55,0,0" VerticalAlignment="Top" Width="49" Grid.ColumnSpan="2"/>
        <TextBox x:Name="txtTo" HorizontalAlignment="Left" Height="27" Margin="60,55,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="423" Grid.ColumnSpan="2"/>
        <Label x:Name="lblCc" Content="CC:" HorizontalAlignment="Left" Height="27" Margin="23,87,0,0" VerticalAlignment="Top" Width="49" Grid.ColumnSpan="2"/>
        <TextBox x:Name="txtCc" HorizontalAlignment="Left" Height="27" Margin="60,87,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="423" Grid.ColumnSpan="2"/>
        <TextBox x:Name="txtBody" HorizontalAlignment="Left" Height="136" Margin="14.088,146,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="414" Grid.Column="1"/>
        <TextBox x:Name="txtSubject" HorizontalAlignment="Left" Height="27" Margin="14.088,117,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="414" Grid.Column="1"/>
        <Label x:Name="lblSubject" Content="Subject" HorizontalAlignment="Left" Height="27" Margin="23,119,0,0" VerticalAlignment="Top" Width="49" Grid.ColumnSpan="2"/>
        <Label x:Name="lblBody" Content="Body" HorizontalAlignment="Left" Height="27" Margin="23,151,0,0" VerticalAlignment="Top" Width="49" Grid.ColumnSpan="2"/>
        <Button x:Name="btnCancel" Content="Cancelar" HorizontalAlignment="Left" Height="31" Margin="223.088,287,0,0" VerticalAlignment="Top" Width="97" Grid.Column="1"/>
        <Button x:Name="btnSend" Content="Enviar" HorizontalAlignment="Left" Height="31" Margin="325.088,287,0,0" VerticalAlignment="Top" Width="98" Grid.Column="1"/>

    </Grid>
</Window>
'@

$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$Form=[Windows.Markup.XamlReader]::Load( $reader )

# Creación del objeto SmtpClient
$client = New-Object System.Net.Mail.SmtpClient 
# Creación del objeto MailMessage
$message = New-Object System.Net.Mail.MailMessage

function enviar{
#Send-email.ps1
#Script que permite el envio de un e-mail
#variable intermedia.
$vartexto=$Form.FindName('txtFrom')
<#$vartexto2=$Form.FindName('txtTo')#>
<#$textoenviar = $Form.FindName('txtFrom')#>
<#$remitente =  $Form.FindName('txtTo') <#’remitente@host.com’#>
$destinatario = $Form.FindName('txtTo').Text <#’felix.avendano@tecandweb.com’#>
$servidor = ’smtp.1and1.es’

$varasunto = $Form.FindName('txtSubject')  <#’Envío de correo vía powershell’ + $(get-date) #>
$varbody=$Form.FindName("txtBody") 
<#$texto = $Form.FindName("txtBody").Text <#’este es el cuerpo del mensaje’ #>






# Adición de las propiedades  
$message.Body = $varbody.Text
$message.From = $vartexto.Text
$message.Subject = $varasunto.Text
$message.To.Add($destinatario)

#credenciales

$secpasswd = ConvertTo-SecureString "colon1001" -AsPlainText -Force

$credenciales = New-Object System.Management.Automation.PSCredential ("alumnos@lhdevsolution.es", $secpasswd)
# ó
#$credenciales = Get-Credential


# Definición de la propiedad del servidor  
$client.Set_Host($servidor)
$client.Credentials=$credenciales
}

# Envío del mensaje con el método Send
$btnBoton2 = $Form.FindName('btnSend')
$btnBoton2.Add_Click({
enviar
Write-Host ("contenido de destinatario" + $vartexto.Text +" contenido de from" + $vartexto2.Text)
Write-Host ($vartexto)
$client.Send($message)
	})

$btnBoton1 = $Form.FindName('btnCancel')
$btnBoton1.add_click({

$Form.Close()
})




$Form.ShowDialog()
#
# Script.ps1
#



