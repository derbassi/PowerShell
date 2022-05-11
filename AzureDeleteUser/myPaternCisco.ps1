#$textFile = c:\temp\vsxinternet01CPK.txt

# Looking for the text using a pattern 
$format = "(?<=Peer: )([0-9]+(\.[0-9]+)+)(?=.- )|(?<=Peer: ([0-9]+(\.[0-9]+)+) - )([a-zA-Z_ ]+)|(?<=Methods: )([a-zA-Z_ -1234568]+)|(?<=My TS:   )([0-9]+(\.[0-9]+)+)-?([0-9]+(\.[0-9]+)+)?|(?<=Peer TS: )([0-9]+(\.[0-9]+)+)-?([0-9]+(\.[0-9]+)+)?"
$theDefinitionFile = select-string -Path c:\temp\vsxinternet01CPK.txt -Pattern $format -AllMatches | % { $_.Matches } | % { $_.Value }

$theDefinitionFile | Out-File ImportToexcel.txt

