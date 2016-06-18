#TFGR

<#.Synopsis

    
    This is a module wich provide some additional funcions used in subscription proccess lacunched by Process-requestsmaster.

.DESCRIPTION
    This is a module wich provide some additional funcions used in subscription proccess lacunched by Process-requestsmaster.
    Funtions:
        get-domainFromUser: Actually it is not needed, but created for scalability porpouses.
        get-OUFromUser:  Actually it is not needed, but created for scalability porpouses. Right now is used just to split Pre-production enviroment users
        get-DetailsFromCommunities: Read Community-ADGroup-ListaMailman list and return all the data in relation with the communities chosen by the user.
.NOTES

    Written by Jorge Martinez, 
    I take no responsibility for any issues caused by this script.

.FUNCTIONALITY
    This is a module wich provide some additional funcions used in subscription proccess lacunched by Process-requestsmaster.

#>


#### FUNCTIONS ########



function get-DetailsFromCommunities
 {
 #TFGR
 <#
    .Synopsis
       Read Community-ADGroup-ListaMailman list and return all the data in relation with the communities chosen by the user.
    
    .EXAMPLE
       
       get-DetailsFromCommunities -communities $spuser.Communities -url $webCommunities -nameList $communityList -userAdmin $userAdmin -sharepointcomputername $sharepointComputerName -Credential $crd
                        
    .OUTPUTS
       [array][array]
    .NOTES
       Written by Jorge Martinez, 
       I take no responsibility for any issues caused by this script.
    .FUNCTIONALITY
       Read Community-ADGroup-ListaMailman list and return all the data in relation with the communities chosen by the user.
       At the momment it returns an array of 4 arrays
        Position 0=All the distribution lists associacted to the communities requested by the user
        Position 1=All the Active driectory groups the user needs to be added 
        Position 2=All the Multi-commuinities the user asked for subcription
        Position 3=All the diferents Support email addresess where the user can ask for help for each different community
#>



 [CmdletBinding(SupportsShouldProcess=$true)]
      param
      (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='¿communities a la que pertence')]
        $Communities,
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='¿url de la lista donde estan la relacion entre communities, AdGroup y mailman')]
        $URL,
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='¿nombre de la lista?')]
        $nameList,
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='¿Usuario administrador')]
        $userAdmin,
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Máquna')]
        $sharepointcomputername,
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Máquna')]
        $Credential
        )


#Inicialización de Variables.
  
  $listasDeCorreo=@()
  $communityinLine=@()
  $ADGroupsTef=@()
  $ADGroupsTef=@()
  $ADGroupsTefBecouseMacro=@()
  $supportEmail=@()
  $macro=@()
    

  $listasDeCorreoString=""
  $ADGroupsString=""
  $ADGroupsTefBecouseMacroString=""
  $supportEmailString=""
  $macroString=""
  $columnToSeekBy="Community"
    
  $comunity=@()
  $i=0  
 
# Main


# There are too much different possibilities, from there is NO community checkt up to all of them have been selected.
# Some errors have been found due to 0 communities selected or just 1 community selected, all of them resolved
# With the While expresion with a 2 levels object condition) an additional error was found. When the "clause while" asks for the next element but this element does not exist becaused the previous element was the last one. #RESOLVED 

$tempstring=$Communities+"," # A comma has to be added in order that allways exits a comma even when there is no community or there is just one

if ($tempstring.length -gt 1) { #if there is at least one community, the first one has to be added because "while cluase" used to procces all the communities selected by the user, hop over the first element
    $comunity+=($tempstring.split(",")[0]).substring(0)
    }

while (($tempstring.substring(($tempstring.indexof(","))+1)).contains(",")) # While the next element has a comma, wich means there is at least one more element to proccess..

{
$comunity+=($tempstring.split(",")[1]).substring(0) # the community is added
$tempstring=$tempstring.substring(($tempstring.indexof(","))+1) # the current community is removed from the temporal variable 
}



$comunity|%{ #for all the communities, all the values in the list are readed and added to the different variables.
    $valueToSeekBy=$_
    $communityinLine=Invoke-Command -Authentication Credssp -computername $sharepointcomputername -command {param($URL,$nameList,$valueToSeekBy,$userAdmin) C:\Admon\Scripts\CommunitiesBI\AD-Communities\SharePoint\SP-leelista-communities.ps1 -URL $URL -nameList $nameList -valueToSeekBy $valueToSeekBy -userAdmin $userAdmin} -credential $Credential -ArgumentList $URL,$nameList,$valueToSeekBy,$userAdmin
    $listasDeCorreo+=$communityinLine.Mailman
    $ADGroupsTef+=$communityinLine.GroupTEF
    $ADGroupsTefBecouseMacro+=$communityinLine.GroupMacro #TBD tenerlo en cuanta
    $supportEmail+=$communityinLine.SupportEmail
    $macro+=$communityinLine.Macro
    }

#For all the variables a trasnformation is needed. All the arrays are transformed to strings with semicolon to split the different unic values.
$i=0
while ($i -lt $listasDeCorreo.count) # 
{
$listasDeCorreoString+=$listasDeCorreo[$i]+";"
$ADGroupsString+=$ADGroupsTef[$i]+";"
if ($ADGroupsTefBecouseMacroString -notcontains $ADGroupsTefBecouseMacro[$i]){$ADGroupsTefBecouseMacroString+=$ADGroupsTefBecouseMacro[$i]+";"}
if ($supportEmailString -notcontains $supportEmail[$i]){$supportEmailString+=$supportEmail[$i]+";"}
if ($macroString-notcontains $macro[$i]){$macroString+=$macro[$i]+";"}
$i+=1
}

$ADGroupsString+=$ADGroupsTefBecouseMacroString # Groups because multicommunities are added to groups becuase communities.
$ADGroupsString+="Communities EMPRESA;" #All the users have to be added to this group.
$ADGroupsString=$ADGroupsString.Replace(";;",";") # double semicolons are removed form the string

  return $listasDeCorreoString,$ADGroupsString,$supportEmailString,$macroString
   
}
<#

 function get-AdGroupsFromUser
 {
param
    (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='¿Domain, please?')]
        [string]$domain,

        [Parameter(Mandatory=$false,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='¿email, please?')]
        $Communities,

        [Parameter(Mandatory=$false,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='¿url de la lista donde estan la relacion entre communities, AdGroup y mailman')]
        $url
   
    )
#TBD descomponemos en comminuties, para cada community leemos que grupo le toca y lo añadimos al campo @#integrarlo con listasmailman

if ($domain -eq "@tef.inet") {return "GG_communities;GG_data_science_and_advance_analytics;GG_Bi_Bigdata capabilities;"}
else {return "escenario a evaluar, usuario que no es de TEF en funcion get-AdGroupsFromUser"}


}#funcion.
#>       


function get-domainFromUser
#TBD (ahora mismo no se utiliza pero se ha preparado por si en un futuro es necesaria)
{
#TFGR

[CmdletBinding(SupportsShouldProcess=$true)]
      param
      (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='¿email, please?')]
        
        [string]$email
    )
       
return "@***.***"
}


function get-OUFromUser
{
#TFGR

#Actualente NO es necesaria pues para todas las communities la OU es la misma, pero se ha creado por motivos de escalabilidad
# Diferenciamos para "PREPRODUCCION"
param
    (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='¿Domain, please?')]
        [string]$userDomain,

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='(PRO,PRE,IC')]
        [string]$entorno,

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='¿Macro-community?')]
        $macroCommunity
   
    )
    write-verbose $entorno
    write-verbose $userDomain
#$macroCommunity #grupo de commnities, que no community. por ahora no se usa pero se puede llegar a usar pra diferenciar disitinas communities.
#


 switch ($entorno)
    {
    
    "PRE" { [string]$OU="OU=QAPRE,OU=Communities TEF,DC=***,DC=***"}

    "PRO" { [string]$OU="OU=Communities ,DC=***,DC=***"}
    
    }#end del switch
 return $OU
 
 }


