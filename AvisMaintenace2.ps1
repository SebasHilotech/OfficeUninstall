Add-Type -AssemblyName PresentationCore,PresentationFramework
$MessageIcon = [System.Windows.MessageBoxImage]::Error
$ButtonType = [System.Windows.MessageBoxButton]::OK
$MessageBody = @"

- Maintenance Imminente - 

Sauvegarder imm�diatement tout vos travaux en cours. 

Votre ordinateur entre dans une maintenance dans 15 minutes. 

Tout travaux en cours non-sauvegard�e seront perdu!

Pri�re de laisser votre ordinateur ouvert lors de plage de maintenance (16-21h)

Merci,

"@
$MessageTitle = "Maintenance Planifi�e - 17h00"

$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
