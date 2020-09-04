Add-Type -AssemblyName PresentationCore,PresentationFramework
$MessageIcon = [System.Windows.MessageBoxImage]::Warning
$ButtonType = [System.Windows.MessageBoxButton]::OK
$MessageBody = @"

- Attention!! Attention!! - Maintenance Planifiée !! -  


Une maintenace planifiée est prévu Ã  17h sur votre ordinateur. 


Prière de sauvegarder vos données avant 17h et de conserver votre ordinateur ouvert entre 16h et 20h.

Prendre note que lors de la maintenance votre ordinateur va redémarrer Ã  quelques reprises.

Il est donc fortement recommendée vous sauvergarder tout travaux en cours avant 16h30. 

Merci,

"@
$MessageTitle = "Maintenance Planifiée - 17h00"

$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
