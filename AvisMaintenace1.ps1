Add-Type -AssemblyName PresentationCore,PresentationFramework
$MessageIcon = [System.Windows.MessageBoxImage]::Warning
$ButtonType = [System.Windows.MessageBoxButton]::OK
$MessageBody = @"

- Attention!! Attention!! - Maintenance Planifi�e !! -  


Une maintenace planifi�e est pr�vu à 17h sur votre ordinateur. 


Pri�re de sauvegarder vos donn�es avant 17h et de conserver votre ordinateur ouvert entre 16h et 20h.

Prendre note que lors de la maintenance votre ordinateur va red�marrer à quelques reprises.

Il est donc fortement recommend�e vous sauvergarder tout travaux en cours avant 16h30. 

Merci,

"@
$MessageTitle = "Maintenance Planifi�e - 17h00"

$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
