# 04 APR 2022
# 03:59

[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows')
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$xaml=@'
<?xml version="1.0" encoding ="utf-16"?>
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" FontFamily="Consolas"
Title="[Azure] Delete User v1.0" Height="475" Width="596" WindowStartupLocation="CenterScreen" BorderBrush="Blue" MinWidth="596" MinHeight="475" MaxWidth="596" MaxHeight="475" ScrollViewer.VerticalScrollBarVisibility="Disabled"> 

<Grid Background="#E3E3E3">
        
<Label Name="LabelAccount" Content="Account:" HorizontalAlignment="Left" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold" Width="65" Height="25" Margin="0,34,0,0"/>
<Label Name="LabelSubscriptionName" Content="Subscription Name:" HorizontalAlignment="Left" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold" Width="131" Height="25" Margin="0,64,0,0"/>
<Label Name="LabelTenantId" Content="Tenant ID:" HorizontalAlignment="Left" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold" Width="131" Height="25" Margin="0,94,0,0"/>

<Button Name="buttonConnect" Content="Connect" HorizontalAlignment="Left" Height="22" Margin="0,142,0,0" VerticalAlignment="Top" Width="492" Opacity="0.8" FontWeight="Bold" FontFamily="Consolas">
<Button.Effect>
<DropShadowEffect/>
</Button.Effect>
        </Button>
        <TextBox Name="TextBoxAccount" HorizontalAlignment="Left" Margin="136,38,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="298" Height="18"/>
        <TextBox Name="TextBoxSubscriptionName" HorizontalAlignment="Left" Margin="136,68,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="298" Height="18"/>
        <TextBox Name="TextBoxTenantId" HorizontalAlignment="Left" Margin="82,98,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="352" Height="18"/>
        <Rectangle HorizontalAlignment="Left" Height="127" Margin="0,10,0,0" Stroke="Black" VerticalAlignment="Top" Width="492"/>
        <TextBlock Name="TextBlock" HorizontalAlignment="Left" Height="229" Margin="0,174,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="492" FontFamily="Consolas" Background="#F0F0F0"/>



        <Ellipse Opacity="{DynamicResource EllipseOpacityConnected}" HorizontalAlignment="Left" Height="18" Margin="17,145,0,0" Stroke="Black" VerticalAlignment="Top" Width="18">
            <Ellipse.Fill>
                <LinearGradientBrush EndPoint="0.5,1" StartPoint="0.5,0">
                    <GradientStop Color="#FF5F0906" Offset="1"/>
                    <GradientStop Color="Black" Offset="1"/>
                </LinearGradientBrush>
            </Ellipse.Fill>
        </Ellipse>

       <Ellipse Opacity="{DynamicResource EllipseOpacityConnected}" HorizontalAlignment="Left" Height="18" Margin="308,144,0,0" Stroke="Black" VerticalAlignment="Top" Width="18">
            <Ellipse.Fill>
                <RadialGradientBrush>
                    <GradientStop Color="#FF0B6827" Offset="1"/>
                    <GradientStop Color="Black" Offset="1"/>
                </RadialGradientBrush>
            </Ellipse.Fill>
        </Ellipse>

        <Button Name="ButtonDeleteUser" Content="Delete User" HorizontalAlignment="Left" Height="22" Margin="10,407,0,0" VerticalAlignment="Top" Width="92" Opacity="0.8" FontWeight="Bold" FontFamily="Consolas">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>   


        <Button Name="Close" Content="Close" HorizontalAlignment="Left" Height="22" Margin="410,407,0,0" VerticalAlignment="Top" Width="92" Opacity="0.8" FontWeight="Bold" FontFamily="Consolas">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>   
            
      </Grid>
</Window>
'@

$myReader = (New-Object System.Xml.XmlNodeReader $xaml)
$myForm = [Windows.Markup.XamlReader]::Load($myReader)

$myForm.FindName("Close").add_click({
    $myForm.Close()
})


$myForm.ShowDialog()