$webAppUrl = "http://yoursharepointsite"
$assembly = "Microsoft.SharePoint.WorkflowActions, Version=16.0.0.0, Culture=neutral, PublicKeyToken=null"
$namespace = "Microsoft.SharePoint.WorkflowActions.WithKey"
$typeName = "*"
$authorized = "True"
$webApp = Get-SPWebApplication $webAppUrl

$xmlElement = "<authorizedType Assembly=`"$assembly`" Namespace=`"$namespace`" TypeName=`"$typeName`" Authorized=`"$authorized`" />"

$modification = New-Object Microsoft.SharePoint.Administration.SPWebConfigModification
$modification.Path = "configuration/System.Workflow.ComponentModel.WorkflowCompiler/authorizedTypes/targetFx[@version='v4.0']"
$modification.Name = "authorizedType[@Assembly='$assembly'][@Namespace='$namespace']"
$modification.Owner = "CustomScript"
$modification.Sequence = 0
$modification.Type = [Microsoft.SharePoint.Administration.SPWebConfigModificationType]::EnsureChildNode
$modification.Value = $xmlElement

$webApp.WebConfigModifications.Add($modification)
$webApp.Update()
$webApp.Farm.Services | ? { $_.GetType().Name -eq "SPWebService" } | % { $_.ApplyWebConfigModifications() }

Write-Host "Modification applied successfully to the web.config of the web application: $webAppUrl"
