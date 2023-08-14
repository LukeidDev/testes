$exclude = @("venv", "vacinaBot.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "vacinaBot.zip" -Force