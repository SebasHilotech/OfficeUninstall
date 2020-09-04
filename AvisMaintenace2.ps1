Add-Type -AssemblyName PresentationCore,PresentationFramework
$MessageIcon = [System.Windows.MessageBoxImage]::Error
$ButtonType = [System.Windows.MessageBoxButton]::OK
$MessageBody = @"

- Maintenance Imminente - 

Sauvegarder immédiatement tout vos travaux en cours. 

Votre ordinateur entre dans une maintenance dans 15 minutes. 

Tout travaux en cours non-sauvegardée seront perdu!

Prière de laisser votre ordinateur ouvert lors de plage de maintenance (16-21h)

Merci,

"@
$MessageTitle = "Maintenance Planifiée - 17h00"

$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
