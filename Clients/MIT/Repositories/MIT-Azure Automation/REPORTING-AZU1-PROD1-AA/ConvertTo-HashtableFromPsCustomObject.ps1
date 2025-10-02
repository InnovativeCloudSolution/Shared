# Credit: https://omgdebugging.com/2019/02/25/convert-a-psobject-to-a-hashtable-in-powershell/

param ( 
    [Parameter(  
        Position = 0,   
        Mandatory = $true,   
        ValueFromPipeline = $true,  
        ValueFromPipelineByPropertyName = $true  
    )] [object] $psCustomObject 
)

$output = @{};
$psCustomObject | Get-Member -MemberType *Property | ForEach-Object {
    $output.($_.name) = $psCustomObject.($_.name)
} 

return $output