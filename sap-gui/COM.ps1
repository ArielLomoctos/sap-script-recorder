
#-Begin-----------------------------------------------------------------

  #-Load assembly-------------------------------------------------------
  [Void] [Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
  [Void] [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

  #-Function Create-Object----------------------------------------------
  Function Create-Object([String]$objectName) {
    try {
      New-Object -ComObject $objectName
    } catch {
      [Void] [System.Windows.Forms.MessageBox]::Show(
        "Can't create object", "Important hint", 0)
    }  
  }

  #-Function Get-Object-------------------------------------------------
  Function Get-Object([String]$pathName, [String]$class) {
    [Microsoft.VisualBasic.Interaction]::GetObject($pathName, $class)
  }

  #-Sub Free-Object-----------------------------------------------------
  Function Free-Object([__ComObject]$object) {
    [Void] [System.Runtime.Interopservices.Marshal]::ReleaseComObject($object)
  }

  #-Function Get-Property-----------------------------------------------
  Function Get-Property([__ComObject]$object, [String]$propertyName, 
    $propertyParameter) {
    $objectType = [System.Type]::GetType($object)
    $objectType.InvokeMember($propertyName,
      [System.Reflection.Bindingflags]::GetProperty,
      $null, $object, $propertyParameter)
  }

  #-Sub Set-Property----------------------------------------------------
  Function Set-Property([__ComObject]$object, [String]$propertyName, 
    $propertyValue) {
    $objectType = [System.Type]::GetType($object)
    [Void] $objectType.InvokeMember($propertyName,
      [System.Reflection.Bindingflags]::SetProperty,
      $null, $object, $propertyValue)
  }

  #-Function Invoke-Method----------------------------------------------
  Function Invoke-Method([__ComObject]$object, [String]$methodName, 
    $methodParameters) {
    $objectType = [System.Type]::GetType($object)
    $output = $objectType.InvokeMember($methodName,
      "InvokeMethod", $NULL, $object, $methodParameters)
    if ( $output ) { $output }
  }

#-End-------------------------------------------------------------------
