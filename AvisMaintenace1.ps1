Add-Type -AssemblyName PresentationCore,PresentationFramework
$MessageIcon = [System.Windows.MessageBoxImage]::Warning
$ButtonType = [System.Windows.MessageBoxButton]::OK
$MessageBody = @"

- Attention!! Attention!! - Maintenance Planifiéee !! -  


Une maintenace planifiée est prévu à 17h sur votre ordinateur. 


Priére de sauvegarder vos données avant 17h et de conserver votre ordinateur ouvert entre 16h et 20h.

Prendre note que lors de la maintenance votre ordinateur va redémarrer à quelques reprises.

Il est donc fortement recommendée vous sauvergarder tout travaux en cours avant 16h30. 

Merci,

"@
$MessageTitle = "Maintenance Planifiée - 17h00"

$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
