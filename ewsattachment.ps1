Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"


$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)


$creds = New-Object System.Net.NetworkCredential("aa@aa.com","") 
$service.Credentials = $creds  
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Compiler=$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
namespace Local.ToolkitExtensions.Net.CertificatePolicy{
public class TrustAll : System.Net.ICertificatePolicy {
public TrustAll() {
}
public bool CheckValidationResult(System.Net.ServicePoint sp,
System.Security.Cryptography.X509Certificates.X509Certificate cert,
System.Net.WebRequest req, int problem) {
return true;
}
}
}
'@
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

## We now create an instance of the TrustAll and attach it to the ServicePointManager
$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll


$MailboxName = "aaa@aaa.com"

$service.AutodiscoverUrl($MailboxName,{$true})



$Sfha = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)

$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)   
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)  

$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(10)
$downloadDirectory = "c:\temp"

$findItemsResults = $Inbox.FindItems($Sfha,$ivItemView)
foreach($miMailItems in $findItemsResults.Items){
       $miMailItems.Load()
       foreach($attach in $miMailItems.Attachments ){
if($attach.Name -like "*.pdf"){
             $attach.Load()
             $fiFile = new-object System.IO.FileStream(($downloadDirectory + “\” + $attach.Name.ToString()), [System.IO.FileMode]::Create)
             $fiFile.Write($attach.Content, 0, $attach.Content.Length)
             $fiFile.Close()
             write-host "Downloaded Attachment : " + (($downloadDirectory + “\” + $attach.Name.ToString()))
}
       }
} 

