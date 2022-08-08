#
#   https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic?view=net-5.0

set-strictMode -version latest

function init() {

 #
 # TODO: Should the following assembly be loaded?
 #
   $assembly = [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

   if ($assembly -eq $null) {
       write-host 'Could not load VB assembly'
   }
}

function typeName($obj) {

   if ($obj -eq $null) {
      return 'null'
   }

 #
 # If $obj is a COM object: return the COM type name.
 # Otherwise, return the PowerShell type name:
 #
   if ($obj -is [System.__ComObject]) {
      return [Microsoft.VisualBasic.Information]::TypeName($obj)
   }

   return $obj.GetType().FullName
}

#
#  $var = new-object 'Microsoft.VisualBasic.VariantType[][]' 5,3
#  ubound($var)          # -> 4
#  ubound($var.item(1))  # -> 2
#
function lBound($obj) {
   return [Microsoft.VisualBasic.Information]::LBound($obj)
}

function uBound($obj) {
   return [Microsoft.VisualBasic.Information]::UBound($obj)
}

function varType($obj) {
   return [Microsoft.VisualBasic.Information]::VarType($obj)
}

# isXyz() functions {

function isArray($obj) {
   return [Microsoft.VisualBasic.Information]::IsArray($obj)
}

function isDate($obj) {
   return [Microsoft.VisualBasic.Information]::IsDate($obj)
}

function isDBNull($obj) {
 #
 # isDBNull $null    -> False
 #
   return [Microsoft.VisualBasic.Information]::IsDBNull($obj)
}

function isNothing($obj) {
 #
 # isNothing $null    -> True
 # isNothing  42
 # isNothing $acc
 #
   return [Microsoft.VisualBasic.Information]::IsNothing($obj)
}

function isError($obj) {
   return [Microsoft.VisualBasic.Information]::IsError($obj)
}

function isReference($obj) {
   return [Microsoft.VisualBasic.Information]::IsReference($obj)
}

# }

function appActivate($procName) {
 #
 # TODO: AppActivate() can be invoked with either the application's title (case insensitive, but no partial name) or
 #       the application's process ID (as is done in the following).
 #
   [Microsoft.VisualBasic.Interaction]::AppActivate( (get-process $procName).id )
}

function rgb($red, $green, $blue) {
   [Microsoft.VisualBasic.Information]::RGB($red, $green, $blue)
}

function callByName {

   param (
      [parameter(mandatory=$true )][__ComObject]                     $obj,
      [parameter(mandatory=$true )][string]                          $proc,
      [parameter(mandatory=$true )][Microsoft.VisualBasic.CallType]  $callType,  # get(2), let(4), set(8), method(1)
      [parameter(mandatory=$false)][object[]]                        $args
   )

   try {
      return [Microsoft.VisualBasic.Interaction]::CallByName($obj, $proc, $callType, $args)
   }
   catch [System.Management.Automation.MethodInvocationException] {
     "callByName: MethodInvocationExceptionException"
      $_ | select *
   }
   catch {
     "callByName: other Exception $($_.GetType().FullName)"
      $_ | select *
   }
}

init
