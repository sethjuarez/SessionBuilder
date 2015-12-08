properties {
	$majorVersion = "1.0"
	$majorWithReleaseVersion = "1.0.0"
	$version = GetVersion $majorWithReleaseVersion
	$treatWarningsAsErrors = $true
	$workingName = if ($workingName) {$workingName} else {"Working"}
	$baseDir  = resolve-path ..
	$buildDir = "$baseDir\Build"
	$sourceDir = "$baseDir\Src"
	$toolsDir = "$baseDir\Tools"
	$workingDir = "$baseDir\$workingName"
	$workingSourceDir = "$workingDir\Src"
	$builds = @(
		@{Name = "numl"; TestsName = "numl.Tests"; BuildFunction = "MSBuildBuild"; TestsFunction = "NUnitTests"; Constants=""; FinalDir="Portable"; NuGetDir = "portable-net45+wp80+win8+wpa81+dnxcore50"; Framework="net-4.0"}
	)
}

framework '4.6x86'

task default -depends Nuget

task Clean {
	Write-Host "Setting location to $baseDir"
	Set-Location $baseDir
	
	if (Test-Path -path $workingDir)
	{
		Write-Host "Deleting existing working directory $workingDir"
	
		Execute-Command -command { del $workingDir -Recurse -Force }
	}
	
	Write-Host "Creating working directory $workingDir"
	New-Item -Path $workingDir -ItemType Directory
}

task Build -depends Clean {
	Write-Host "Copying source to working source directory $workingSourceDir"
	robocopy $sourceDir $workingSourceDir /MIR /NP /XD bin obj TestResults packages $packageDirs .vs artifacts /XF *.suo *.user *.lock.json | Out-Default
	
	Write-Host -ForegroundColor Green "Updating assembly version"
	Write-Host
	Update-AssemblyInfoFiles $workingSourceDir ($majorVersion + '.0.0') $version
	
	Update-Project $workingSourceDir\numl\project.json
	
	foreach ($build in $builds)
	{
		$name = $build.Name
		if ($name -ne $null)
		{
			Write-Host -ForegroundColor Green "Building " $name
			& $build.BuildFunction $build
		}
	}
}

task ILMerge -depends Build {
	
}

function Update-AssemblyInfoFiles ([string] $workingSourceDir, [string] $assemblyVersionNumber, [string] $fileVersionNumber)
{
	$assemblyVersionPattern = 'AssemblyVersion\("[0-9]+(\.([0-9]+|\*)){1,3}"\)'
	$fileVersionPattern = 'AssemblyFileVersion\("[0-9]+(\.([0-9]+|\*)){1,3}"\)'
	$assemblyVersion = 'AssemblyVersion("' + $assemblyVersionNumber + '")';
	$fileVersion = 'AssemblyFileVersion("' + $fileVersionNumber + '")';
	
	Get-ChildItem -Path $workingSourceDir -r -filter AssemblyInfo.cs | ForEach-Object {
		Write-Host $workingSourceDir
		$filename = $_.Directory.ToString() + '\' + $_.Name
		Write-Host $filename
		$filename + ' -> ' + $version
		
		(Get-Content $filename) | ForEach-Object {
			% {$_ -replace $assemblyVersionPattern, $assemblyVersion } |
			% {$_ -replace $fileVersionPattern, $fileVersion }
			
		} | Set-Content $filename
	}
}

function Update-Project($projectPath) 
{
	$json = (Get-Content $projectPath) -join "`n" | ConvertFrom-Json
	$options = @{"warningsAsErrors" = $true; "define" = ((GetConstants "dotnet") -split ";") }
	Add-Member -InputObject $json -MemberType NoteProperty -Name "compilationOptions" -Value $options -Force
	ConvertTo-Json $json -Depth 10 | Set-Content $projectPath
}

function GetConstants($constants)
{
	return "CODE_ANALYSIS;TRACE;$constants"
}

function MSBuildBuild($build)
{
	$name = $build.Name
	$finalDir = $build.FinalDir
	
	Write-Host
	Write-Host "Restoring $workingSourceDir\$name.sln" -ForegroundColor Green
	[Environment]::SetEnvironmentVariable("EnableNuGetPackageRestore", "true", "Process")
	exec { .\Tools\NuGet\NuGet.exe update -self }
	exec { .\Tools\NuGet\NuGet.exe restore "$workingSourceDir\$name.sln" -verbosity detailed -configfile $workingSourceDir\nuget.config | Out-Default } "Error restoring $name"
	
	$constants = GetConstants $build.Constants
	
	Write-Host
	Write-Host "Building $workingSourceDir\$name.sln" -ForegroundColor Green
	if ($name -match 'Tests') { 
		Write-Host "NO DOCS $name" -ForegroundColor Red
		#exec { msbuild "/t:Clean;Rebuild" /p:Configuration=Release /p:OutputPath=bin\Release\$finalDir\ "/p:TreatWarningsAsErrors=$treatWarningsAsErrors" "/p:VisualStudioVersion=14.0" /p:DefineConstants=`"$constants`" "$workingSourceDir\$name.sln" | Out-Default } "Error building $name"
	} else {
		Write-Host "YES DOCS $name" -ForegroundColor Red
		#exec { msbuild "/t:Clean;Rebuild" /p:Configuration=Release /p:OutputPath=bin\Release\$finalDir\ "/p:TreatWarningsAsErrors=$treatWarningsAsErrors" "/p:VisualStudioVersion=14.0" "/p:documentationFile=bin\Release\$finalDir\$name.xml" /p:DefineConstants=`"$constants`" "$workingSourceDir\$name.sln" | Out-Default } "Error building $name"	
	}
}

function GetVersion($majorVersion)
{
	$now = [DateTime]::Now
	
	$year = $now.Year - 2000
	$month = $now.Month
	$totalMonthsSince2000 = ($year * 12) + $month
	$day = $now.Day
	$minor = "{0}{1:00}" -f $totalMonthsSince2000, $day
	
	$hour = $now.Hour
	$minute = $now.Minute
	$revision = "{0:00}{1:00}" -f $hour, $minute
	
	return $majorVersion + "." + $minor
}

function Edit-XmlNodes {
    param (
        [xml] $doc,
        [string] $xpath = $(throw "xpath is a required parameter"),
        [string] $value = $(throw "value is a required parameter")
    )
    
    $nodes = $doc.SelectNodes($xpath)
    $count = $nodes.Count

    Write-Host "Found $count nodes with path '$xpath'"
    
    foreach ($node in $nodes) {
        if ($node -ne $null) {
            if ($node.NodeType -eq "Element")
            {
                $node.InnerXml = $value
            }
            else
            {
                $node.Value = $value
            }
        }
    }
}

function Execute-Command($command) 
{
	$currentRetry = 0
	$success = $false
	do 
	{
		try
		{
			& $command
			$success = $true
		}
		catch [System.Exception]
		{
			if ($currentRetry -gt 5) 
			{
				throw $_.Exception.ToString()
			} 
			else 
			{
				write-host "Retry $currentRetry"
				Start-Sleep -s 1
			}
			$currentRetry = $currentRetry + 1
		}
	} 
	while (!$success)
}