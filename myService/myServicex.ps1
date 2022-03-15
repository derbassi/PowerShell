Add-Type -AssemblyName presentationframework
[xml]$xaml=@'
<?xml version="1.0" encoding ="utf-16"?>
<Window 
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
Name="myService" Width="500" MinHeight="462" MinWidth="588" ResizeMode="NoResize" SizeToContent="Height" Title="myService v1.0" WindowStartupLocation="CenterScreen" Height="460.667">
    <Grid Margin="0,0,4.667,21">
        <Rectangle Fill="#FFF4F4F5" HorizontalAlignment="Left" Height="78" Margin="19,64,0,0" Stroke="Black" VerticalAlignment="Top" Width="545"/>
        <Label Content="Service Name:" Width="107" Height="26" Margin="20,22,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Grid.Column="0" Grid.Row= "0" FontFamily="Consolas" FontWeight="Bold" />
        <TextBox  Width="170" Height="22" Margin="127,26,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Grid.Column="0" Grid.Row="0" />
        <TextBox Width="544" Height="239" Margin="20,157,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" HorizontalContentAlignment="Left" HorizontalScrollBarVisibility="Auto" IsReadOnly="True"  Text="" VerticalScrollBarVisibility="Auto" Grid.Column="0" Grid.Row="0" FontFamily="Consolas" FontSize="10" />
        <Button Name="Go" Content="Go" Width="241" Height="20" Margin="0,0,13.333,364" HorizontalAlignment="Right" VerticalAlignment="Bottom" Grid.Column="0" Grid.Row="0" RenderTransformOrigin="-0.796,0.283" FontFamily="Consolas" />
        <CheckBox  Content="Stop" HorizontalAlignment="Left" Margin="34,80,0,0" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold"/>
        <CheckBox Content="Start" HorizontalAlignment="Left" Margin="127,80,0,0" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold"/>
        <ListBox HorizontalAlignment="Left" Height="55" Margin="448,77,0,0" VerticalAlignment="Top" Width="100" FontFamily="Consolas" FontSize="10">
            <ListBoxItem Content="Auto" Background="#FFB3DC5B"/>
            <ListBoxItem Content="Manual" Background="#FFB0B5D4"/>
            <ListBoxItem Content="Disabled" Background="#FFEEB9B9"/>
        </ListBox>
        <CheckBox Content="Process ID" HorizontalAlignment="Left" Margin="34,117,0,0" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold"/>
        <CheckBox Content="Required Service" HorizontalAlignment="Left" Margin="270,80,0,0" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold" RenderTransformOrigin="0.472,0.267"/>
        <CheckBox Content="Dependancy Service" HorizontalAlignment="Left" Margin="270,117,0,0" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold" RenderTransformOrigin="0.472,0.267"/>
        <Button  Name="Close" Content="Close" Width="75" Height="20" Margin="0,0,483,-16" HorizontalAlignment="Right" VerticalAlignment="Bottom" Grid.Column="0" Grid.Row="0" />

    </Grid>
</Window>
'@


$myReader = (New-Object System.Xml.XmlNodeReader $xaml)
$myForm = [Windows.Markup.XamlReader]::Load($myReader)

$myForm.FindName("Close").add_click({
    $myForm.Close()
})

=======
﻿Add-Type -AssemblyName presentationframework
[xml]$xaml=@'
<?xml version="1.0" encoding ="utf-16"?>
<Window 
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
Name="myService" Width="500" MinHeight="462" MinWidth="588" ResizeMode="NoResize" SizeToContent="Height" Title="myService v1.0" WindowStartupLocation="CenterScreen" Height="460.667">
    <Grid Margin="0,0,4.667,21">
        <Rectangle Fill="#FFF4F4F5" HorizontalAlignment="Left" Height="78" Margin="19,64,0,0" Stroke="Black" VerticalAlignment="Top" Width="545"/>
        <Label Content="Service Name:" Width="107" Height="26" Margin="20,22,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Grid.Column="0" Grid.Row= "0" FontFamily="Consolas" FontWeight="Bold" />
        <TextBox  Width="170" Height="22" Margin="127,26,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Grid.Column="0" Grid.Row="0" />
        <TextBox Width="544" Height="239" Margin="20,157,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" HorizontalContentAlignment="Left" HorizontalScrollBarVisibility="Auto" IsReadOnly="True"  Text="" VerticalScrollBarVisibility="Auto" Grid.Column="0" Grid.Row="0" FontFamily="Consolas" FontSize="10" />
        <Button Name="Go" Content="Go" Width="241" Height="20" Margin="0,0,13.333,364" HorizontalAlignment="Right" VerticalAlignment="Bottom" Grid.Column="0" Grid.Row="0" RenderTransformOrigin="-0.796,0.283" FontFamily="Consolas" />
        <CheckBox  Content="Stop" HorizontalAlignment="Left" Margin="34,80,0,0" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold"/>
        <CheckBox Content="Start" HorizontalAlignment="Left" Margin="127,80,0,0" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold"/>
        <ListBox HorizontalAlignment="Left" Height="55" Margin="448,77,0,0" VerticalAlignment="Top" Width="100" FontFamily="Consolas" FontSize="10">
            <ListBoxItem Content="Auto" Background="#FFB3DC5B"/>
            <ListBoxItem Content="Manual" Background="#FFB0B5D4"/>
            <ListBoxItem Content="Disabled" Background="#FFEEB9B9"/>
        </ListBox>
        <CheckBox Content="Process ID" HorizontalAlignment="Left" Margin="34,117,0,0" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold"/>
        <CheckBox Content="Required Service" HorizontalAlignment="Left" Margin="270,80,0,0" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold" RenderTransformOrigin="0.472,0.267"/>
        <CheckBox Content="Dependancy Service" HorizontalAlignment="Left" Margin="270,117,0,0" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold" RenderTransformOrigin="0.472,0.267"/>
        <Button  Name="Close" Content="Close" Width="75" Height="20" Margin="0,0,483,-16" HorizontalAlignment="Right" VerticalAlignment="Bottom" Grid.Column="0" Grid.Row="0" />

    </Grid>
</Window>
'@


$myReader = (New-Object System.Xml.XmlNodeReader $xaml)
$myForm = [Windows.Markup.XamlReader]::Load($myReader)

$myForm.FindName("Close").add_click({
    $myForm.Close()
})

>>>>>>> 20ffb7ea676d35e1a9e97efd0a21efc72fc9c538
$myForm.ShowDialog()
