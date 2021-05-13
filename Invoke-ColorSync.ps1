
<#
.NOTES
    Name: Invoke-ColorSync.ps1
    Original Author: herrlukas
    Author: herrlukas
    Contributor:
    Requires: Razer Chroma SDK
    Major Release History:
    5/15/2021 - Initial Public Release.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
.SYNOPSIS
    Syncs the Windows accent color to any attached Razer keyboard
.DESCRIPTION
    This script syncs the Windows accent color to any attached Razer keyboard using the
    Razer Chroma Rest API.
#>

###############################################################################
# Main Variables                                                              #
###############################################################################

$global:uri = "http://localhost:54235/razer/chromasdk"
$global:colorBGR_old = ""
$global:keyboardEndpoint = "keyboard"
$global:heartbeatEndpoint = "heartbeat"

function Open-ChromaSDK {
    $body = '{
        "title": "Sync-WinAccentToRazer",
        "description": "Tool for syncing the Windows Accent Color to Razer Chroma.",
        "author": {
            "name": "herr_lukas_",
            "contact": "lukas@herrlukas.de"
        },
        "device_supported": [
            "keyboard"
        ],
        "category": "application"
    }'
 
    $parameters = @{
        Method      = "POST"
        Uri         = $uri
        Body        = $body
        ContentType = "application/json"
    }
    
    $response = Invoke-WebRequest @parameters
    
    $response = $response | ConvertFrom-Json
    $global:uri = $response.uri + "/"
}

###############################################################################
# Functions                                                                   #
###############################################################################

function Close-ChromaSDK {
    $parameters = @{
        Method      = 'DELETE'
        Uri         = $global:uri
        ContentType = 'application/json'
    }
    
    Invoke-RestMethod @parameters | Out-Null
}

function Get-AccentColor {
    $colorABGR = (Get-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\DWM).AccentColor
    $global:colorBGR = 0xff000000 -bXOR $colorABGR
}

function Set-KeyboardColor {
    $body = @"
{
    "effect": "CHROMA_STATIC",
    "param": {
        "color": $($colorBGR)
    }
}
"@

    $parameters = @{
        Method      = "PUT"
        Uri         = $global:uri + $global:keyboardEndpoint
        Body        = $body
        ContentType = "application/json"
    }
    Invoke-WebRequest @parameters | Out-Null
}

function Invoke-Heartbeat {
    $parameters = @{
        Method      = 'PUT'
        Uri         = $global:uri + $global:heartbeatEndpoint
        ContentType = 'application/json'
    }
    Invoke-RestMethod @parameters | Out-Null
}

###############################################################################
# Main Block                                                                  #
###############################################################################

Open-ChromaSDK
#Wait for the SDK to start
Start-Sleep -s 3

try {
    while ($true) {
        Get-AccentColor
        if ($global:colorBGR -ne $global:colorBGR_old) {
            Set-KeyboardColor
            $global:colorBGR_old = $global:colorBGR
        }
        else {
            Invoke-Heartbeat
        }
        Start-Sleep -s 1
    }
}
finally {
    Close-ChromaSDK
}
