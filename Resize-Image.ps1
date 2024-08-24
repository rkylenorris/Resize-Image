
<#PSScriptInfo

.VERSION 1.0

.GUID 0cb67c97-1a2f-44f3-a85e-6077e4e06767

.AUTHOR roder

.COMPANYNAME

.COPYRIGHT

.TAGS

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES 
ResizeImageModule
.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

<# 

.DESCRIPTION 
 A simple shell around ResizeImageModule to resize images using a proporitonal multiplier 

#> 
param(
    # path to image to be resized
    [Paramater(Mandatory)]
    [string]
    [ValidatePattern("[.].[bmp|gif|jpeg|jpg|png|tif|tiff]")]
    $ImagePath,
    # decimal proportion for dimenional resizing
    [parameter(Mandatory)]
    [ValidateRange(0,1)]
    [decimal]
    $multiplier,
    # Option to specify path to save to
    [Parameter(Mandatory=$false)]
    [string]
    $ResultsPath
)

function Get-RequiredModule {
    [CmdletBinding()]
    param(
            [Parameter(Mandatory)]
            [string]$ModuleName
        )
    begin {
        
    }
    
    process {
        if (Get-Module | Where-Object { $_.Name -eq $ModuleName }) {
            return $true
        }
        else {
            if (Get-Module -ListAvailable | Where-Object { $_.Name -eq $ModuleName }) {
                Import-Module $ModuleName
                return $true
            }
            else {
                if (Find-Module -Name $ModuleName | Where-Object { $_.Name -eq $ModuleName }) {
                    Install-Module -Name $ModuleName -Force -Scope CurrentUser
                    Import-Module $ModuleName
                    return $true
                }
                else {
                    return $false
                }
            }
        }
    }
    
    end {
        
    }
}

# import module
$moduleLoaded = Get-RequiredModule "ResizeImageModule"

if(-not($moduleLoaded)){
    Write-Warning "Unable to load necessary module; terminating..."
    Read-Host "press enter to exit"
    exit
}

# load image to get dimensions
$image = New-Object -ComObject Wia.ImageFile
$image.LoadFile($ImagePath)
$ext = $image.FileExtension

# create proportional dimensions using multiplier
$imageDims = $($image.Width, $image.Height | ForEach-Object {
    $newDim = $_ * $multiplier
    return $newDim
})

# create default results path if none supplied
if(-not($ResultsPath)){
    $ResultsPath = "./Resized_$([math]::Round($multiplier * 100))_$([System.IO.Path]::GetFileNameWithoutExtension($ImagePath))$ext"
}

# call resize image
try{
    Resize-Image -InputFile $ImagePath -OutputFile $ResultsPath -Width $imageDims[0] -Height $imageDims[1] -ProportionalResize $true\
    Write-Output "Image resized, saved to $ResultsPath"
}catch{
    Write-Warning "Unable to resize image:"
    Write-Warning $_.ErrorDetails
    throw $_.Exception
}


