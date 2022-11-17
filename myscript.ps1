Write-Host "Welcome to demo of powershell prompt input" -ForegroundColor Green
$path= Read-Host -Prompt "Enter your Path to directory"
Write-Host "The entered name is" $path -ForegroundColor Green
$files = Get-ChildItem $path

foreach ($f in $files){
	write-Host $f.FullName
	$wd = new-object -comobject word.application
	$wd.documents.open($f.FullName)
	$wd.run("Tables")
}


