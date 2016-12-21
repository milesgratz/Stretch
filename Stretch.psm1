# LOAD [Microsoft.Toolkit.Uwp.Notifications] LIBRARY
Try { Add-Type -Path $PSScriptRoot\Library\Microsoft.Toolkit.Uwp.Notifications\Microsoft.Toolkit.Uwp.Notifications.dll }
Catch { Write-Error -Message "Failed to load [Microsoft.Toolkit.Uwp.Notifications] library"; Exit }

# BEGIN STRETCH NOTIFICATION FUNCTION
function New-StretchNotification {
    <#
        .SYNOPSIS
        Improve physical health (stretching) by displaying Windows 10 reminder Notifications
        
        .DESCRIPTION
        The New-StretchNotification cmdlet creates and displays a Notification on Windows 10 operating systems.
        You can optionally call the New-StretchNotification cmdlet with the Stretch alias.

        .EXAMPLE
        New-StretchNotification
        
        This command creates and displays a Notification to stretch (random body part.)
        
        .EXAMPLE
        New-StretchNotification -Body Wrist

        This command creates and displays a Notification to stretch your wrists.
        
        .LINK
        https://github.com/milesgratz/Stretch
    #>
	[alias('Stretch')]
    param
    (
        [Parameter(ParameterSetName = 'Body')]
        [ValidateSet('Default',
                    'Ankle',
                    'Chest',
                    'Elbow',
                    'Hip',
                    'Knee', 
                    'Leg',
                    'LowerBack',
                    'Neck',
                    'Shoulders',
                    'UpperBack',
                    'Wrist')]
        [String] $Body = 'Default'
    )

    # DEFINE TOAST LOGO
    $ImagePath = "$PSScriptRoot\Images\Stretch.png"
    $Image = [Microsoft.Toolkit.Uwp.Notifications.ToastGenericAppLogo]::new()
    $Image.Source = $ImagePath
    $Image.HintCrop = "Circle"
	
    # DETERMINE TYPES OF STRETCH
    $Ankle = "Stretch","Take a break and stretch your ankles!","https://www.youtube.com/watch?v=TlHWnNiYjuw"
    $Chest = "Stretch","Take a break and stretch your chest!","https://www.youtube.com/watch?v=MPnAlU70e7U"
    $Elbow = "Stretch","Take a break and stretch your elbows!","https://www.youtube.com/watch?v=POqVlGFfbfk"
    $Hip = "Stretch","Take a break and stretch your hips!","https://www.youtube.com/watch?v=HhF5r8IiX_k"
    $Knee = "Stretch","Take a break and stretch your knees!","https://www.youtube.com/watch?v=YdpiLd2Zx9U"
    $Leg = "Stretch","Take a break and stretch your legs!","https://www.youtube.com/watch?v=JBqu7Xjz1uk"
    $LowerBack = "Stretch","Take a break and stretch your lower back!","https://www.youtube.com/watch?v=jaji1zuVAQU"
    $Neck = "Stretch","Take a break and stretch your neck!","https://www.youtube.com/watch?v=5EvVD-b_b_o"
    $Shoulder = "Stretch","Take a break and stretch your shoulders!","https://www.youtube.com/watch?v=zcriNQ9D3dQ"
    $UpperBack = "Stretch","Take a break and stretch your upper back!","https://www.youtube.com/watch?v=5EvVD-b_b_o"
    $Wrist = "Stretch","Take a break and stretch your wrists!","https://www.youtube.com/watch?v=iL59TI2gihI"

    # DEFAULT STRETCH
    If ($Body -eq "Default")
    {
        $Stretch = Get-Random ($Ankle, $Chest, $Elbow, $Hip, $Knee, $Leg, $LowerBack, $Neck, $Shoulder, $UpperBack, $Wrist)
        $Text = $Stretch[0..1]
    }
    # SPECIFIC STRETCH BASED ON -BODY PARAMETER
    Else
    {
        $Stretch =  ($Ankle, $Chest, $Elbow, $Hip, $Knee, $Leg, $LowerBack, $Neck, $Shoulder, $UpperBack, $Wrist) | Where-Object { $_ -match $Body }
        $Text = $Stretch[0..1]
    }

    # CONVERT PLAINTEXT TO NEW TOAST ADAPTIVETEXT OBJECT
	$TextObjects = @()
    foreach ($Txt in $Text)
    {
		$TextObj = [Microsoft.Toolkit.Uwp.Notifications.AdaptiveText]::new()
        $TextObj.Text = $Txt
		$TextObjects += $TextObj
    }
	
    # ATTACH TEXT OBJECTS TO NEW GENERIC TOAST OBJECT
    $Binding = [Microsoft.Toolkit.Uwp.Notifications.ToastBindingGeneric]::new()
    $Binding.AppLogoOverride = $Image
	foreach ($TextObject in $TextObjects)
	{
		$Binding.Children.Add($TextObject)
	}
    
    # ATTACH NEW GENERIC TOAST OBJECT TO TOAST VISUAL OBJECT
    $Visual = [Microsoft.Toolkit.Uwp.Notifications.ToastVisual]::new()
    $Visual.BindingGeneric = $Binding
    
    # DEFINE TOAST ACTIONS
    $ToastActions = [Microsoft.Toolkit.Uwp.Notifications.ToastActionsCustom]::new()
    $SuggestionsButton = [Microsoft.Toolkit.Uwp.Notifications.ToastButton]::new("Suggestions",$Stretch[2])
    $SuggestionsButton.ActivationType = "protocol"
    $DismissButton = [Microsoft.Toolkit.Uwp.Notifications.ToastButtonDismiss]::new()
    $ToastActions.Buttons.Add($SuggestionsButton)
    $ToastActions.Buttons.Add($DismissButton)

    # ATTACH TOAST VISUAL TO TOAST CONTENT OBJECT
    $ToastContent = [Microsoft.Toolkit.Uwp.Notifications.ToastContent]::new()
    $ToastContent.Actions = $ToastActions
    $ToastContent.Visual = $Visual
    
    # LOAD CORRECT LIBRARY (Microsoft.Toolkit.Uwp.Notifications.dll)
    $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
    $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]

    # CONVERT TOAST CONTENT OBJECT TO XML
    $ToastXml = [Windows.Data.Xml.Dom.XmlDocument]::new()
    $ToastXml.LoadXml($ToastContent.GetContent())

    # CONVERT TOAST XML TO TOAST NOTIFICATION
    $Toast = [Windows.UI.Notifications.ToastNotification]::new($ToastXml)

    # DISPLAY TOAST NOTIFICATION
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Stretch").Show($Toast)
}

# EXPORT STRETCH ALIAS 
Export-ModuleMember -Function New-StretchNotification -Alias Stretch