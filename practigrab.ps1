# MC Hammer time super match script for USPSA
# 2015-04-28
# To do:
# - Smarter stage name extraction
#
# What does she do?
# Given the practiscore URL, shooter name and division the script will get the results for each stage and spit out a series of files
# culminating in a text file that is ready to paste into match videos.

param ($URI = "https://practiscore.com/results.php?uuid=91103c8e-d09b-4f13-9dcb-95024f0081b5", $USPSANumber = "A85001", [switch]$ClassCalc, [switch]$Verbose, [switch]$Debug)


# Logging function
function logAndWrite ($logPath, $logText) {
	#Write-Host $logText
	$logText = ((Get-Date).ToString("yyyyMMdd-hh:mm:ss")) + ": $logText"
	Add-Content -Path $logPath -Value $logText -Force
}

# Creates the specified folder if it does not already exist.
function createFolder ($folderPath) {
	if (!(Test-Path $folderPath)) {
		#Write-Host "Folder path, $folderPath, does not exist. Creating directory..."
		mkdir $folderPath
		
		if (!(Test-Path $folderPath)) {
			#Write-Host "Directory creation failed. Exiting script..."
			Exit 1
		} else {
			#Write-Host "Directory creation succeeded..."
		}
	}
}

# Practiscore URL formats seem to change with the wind. There are no bad error codes if you use the wrong case for example. This will make sure the URL is valid.
function Test-PractiscoreURL ($uri)
{
	if ((getHTML $uri).rawcontent.Contains($match_name))
	{
		return $true
	}
	else
	{
		return $false
	}
}

# Get HTML content for a given URI. Handles errors because practiscore like to fail 1 out of 20 times. 5 failures in a row kills the script.
function getHTML ($uri) {
	$count = 0
	$success = $false
	Write-Debug "Function: getHTML; Input: uri; Value: *$uri*"
	while (($count -lt $timeout_num) -and (!$success)) {
		try {
			$html = Invoke-WebRequest -Uri $uri
			$success = $true
		}
		catch {
			$count++
			if ($count -eq 5) {
				Write-Host "Failed to get HTML. Exiting script..."
				Exit 1
			}
			Start-Sleep $timeout_sleep_sec
		}
	}
	
	# Return html object.
	$html
}

# Derive the stage specific URI given the base URI, stage number and division. 
function getStageUri ($uri, $stage_num, $division) {
	
	# Array to hold URI formats.
	$arr_stage_uri = @()
	
	# Get the short name for the division that practiscore uses in their URL.
	$short_div = $hash_division.Get_Item($division)
	
	# The typical URL format I've seen on practiscore
	$num = $stage_num - 1
	$arr_stage_uri += "$($uri)&page=stage$($num)-$($short_div)"
	
	# The other URL format I've seen in the wild
	$num = $stage_num
	$short_div_upper = $short_div.ToUpper()
	$arr_stage_uri += "$($uri)&page=stage$($num)$($short_div_upper)"

	Foreach ($stage_uri in $arr_stage_uri)
	{
		if (Test-PractiscoreURL $stage_uri)
		{
			return $stage_uri
		}
	}
	
	# No URLs returned valid results.
	Write-Host "ERROR: Could not URL for stage $stage_num and division $division using URL, $stage_uri." -ForegroundColor Red -BackgroundColor Black
	Write-Host "Exiting script."
	Exit 1
}

# Derive the stage specifc URI given the base URI, stage number and division. 
function Get-OverAllURI ($base_uri, $division) {

	# Return Overall Combined results
	if ($division -eq $null)
	{	
		$overall_uri = "$($uri)&page=overall-combined"
		if (Test-PractiscoreURL $overall_uri)
		{
			return $overall_uri
		}
		else
		{
			Write-Host "ERROR: Could not find overall combined URL. Exiting script." -ForegroundColor Red -BackgroundColor Black
			Write-Host "Exiting script."
		}
	}
	# Return Division overall results
	else
	{
		$short_div = $hash_division_extended.Get_Item($division)
		$overall_uri = "$($uri)&page=overall-$short_div "
		
		if (Test-PractiscoreURL $overall_uri)
		{
			return $overall_uri
		}
		else
		{
			Write-Host "ERROR: Could not find URL for overall $division. Exiting script." -ForegroundColor Red -BackgroundColor Black
			Write-Host "Exiting script."
		}
	}
}


# Return an array of stage result objects.
function getStageResults($html, $stage_num) {
	# Hash table to store table headers.
	$hash_headers = @{}

	# Store header names in hash table
	$count = 0
	$html.ParsedHTML.getElementsByTagName("thead") | %{
		$_.getElementsByTagName("th") | %{
			$hash_headers.Add($count,$_.innerText)
			$count++
		}
	}
	
	# Get individual shooter results and add them to a report for the stage. This part is VERY slow in x64 PowerShell.
	# Run in x86 PowerShell for best performance.
	$arr_results = @()
	$count = $null
	$html.ParsedHTML.getElementsByTagName("tr") | %{
		if ($_.getElementsByTagName("td") -ne $null) {
			$count = 0
			$obj = $null
			$obj = New-Object PSObject
			Add-Member -InputObject $obj -MemberType NoteProperty -Name "StageNum" -Value $stage_num
			$_.getElementsByTagName("td") | %{
				Add-Member -InputObject $obj -MemberType NoteProperty -Name $hash_headers.Get_Item($count) -Value $_.innerText
				$count++
			}
			$arr_results += $obj
		}
	}
	
	# Return stage results array.
	$arr_results
}

# Return an array of stage result objects.
function Get-Results($html) {
	# Hash table to store table headers.
	$hash_headers = @{}

	# Store header names in hash table
	$count = 0
	$html.ParsedHTML.getElementsByTagName("thead") | %{
		$_.getElementsByTagName("th") | %{
			$hash_headers.Add($count,$_.innerText)
			$count++
		}
	}
	
	# Get individual shooter results and add them to a report for the stage. This part is VERY slow in x64 PowerShell.
	# Run in x86 PowerShell for best performance.
	$arr_results = @()
	$count = $null
	$html.ParsedHTML.getElementsByTagName("tr") | %{
		if ($_.getElementsByTagName("td") -ne $null) {
			$count = 0
			$obj = $null
			$obj = New-Object PSObject
			$_.getElementsByTagName("td") | %{
				Add-Member -InputObject $obj -MemberType NoteProperty -Name $hash_headers.Get_Item($count) -Value $_.innerText
				$count++
			}
			$arr_results += $obj
		}
	}
	
	# Return stage results array.
	$arr_results
}

# Determine how many stages and what their names are. Store them in hash table and return it.
function getStageName ($html) {
	$hash_stage_names = @{}
	$count = 0
	$html.ParsedHTML.getElementsByTagName("tr") | %{
		if ($_.getElementsByTagName("td") -ne $null) {
			$_.getElementsByTagName("td") | %{
				# There are multiple table elements. We want the ones that contain the word 'stage'.
				if ($_.innerText.Contains("Stage")) {
					$count++
					$stage_name_full = $_.innerText
					
					# This logic can be improved.
					$stage_name = $stage_name_full.Replace("Stage $count ", "").Replace("Bay $count - ", "")
					$hash_stage_names.Add($count,$stage_name)
				}
			}
		}
	}
	
	# Return the hash of stage names. Of format [int]StageNum,[string]"StageName"
	$hash_stage_names
}

function Get-MatchName ($uri)
{
	$html = getHTML $uri
	($html.ParsedHTML.getElementsByTagName("td") | select innertext)[1].innerText
}

function Get-ShooterName ($overall_uri)
{
	$arr_results = Get-Results (getHTML $overall_uri)
	$name = ($arr_results | Where {$_.USPSA -eq $USPSANumber}).Name
	$name
}

function Get-ShooterDivision ($overall_uri)
{
	$arr_results = Get-Results (getHTML $overall_uri)
	$division = ($arr_results | Where {$_.USPSA -eq $USPSANumber}).Division
	$division
}

function Get-ShooterClass ($overall_uri)
{
	$arr_results = Get-Results (getHTML $overall_uri)
	$class = ($arr_results | Where {$_.USPSA -eq $USPSANumber}).Class
	$class
}

function Get-ClassPlace ($arr_class_place, $shooter_name)
{
	
	$arr_class_place = $arr_class_place | Sort "Hit Factor" -Descending
	$place = 0
	foreach ($shooter in $arr_class_place)
	{
		$place++
		if ($arr_class_place.StageNum -eq "3") {$arr_class_place | export-csv c:\temp\dumb$place.csv -notype}
		if ($shooter.Name -eq $shooter_name)
		{
			#Write-Host "place: $place"
			return $place
		}
	}
	
	Write-Host "ERROR: Shooter not found in class results." -ForegroundColor Red -BackgroundColor Black
	Write-Host "Exiting script..."
	Exit 1
}

function Convert-HitFactorToNum ($arr_shooters)
{
	$new_shooter = @()
	foreach ($shooter in $arr_shooters)
	{
		$shooter."Hit Factor" = $shooter."Hit Factor" -as [decimal]
		$new_shooter += $shooter
	}
	$new_shooter
}

# Determine the correct place 'ending' based on the last character of the place. (e.g. input = '23', output = "23rd")
function getPlaceFull ([string]$place) {
	# Get last char of place
	[string]$last_char = $place.SubString($place.Length-1)
	$last_two_char = $null
	if ($place.Length -ge 2)
	{
		[string]$last_two_char = $place.SubString($place.Length-2)
	}
	
	switch ($last_char)
	{
		"1" {
				if ($last_two_char -eq "11")
				{
					$place_end = "th"
				}
				else
				{
					$place_end = "st"
				}
			}
		"2" {
				if ($last_two_char -eq "12")
				{
					$place_end = "th"
				}
				else
				{
					$place_end = "nd"
				}
			}
		"3" {
				if ($last_two_char -eq "13")
				{
					$place_end = "th"
				}
				else
				{
					$place_end = "rd"
				}
			}
		default { $place_end = "th" }
	}
	
	$place_full = "$($place)$($place_end)"
	$place_full
}

function GenerateOutputWithFormat ($obj, $format)
{
	$string = $format
	$string = $string.Replace("<stagenum>", $obj.StageNum)
	$string = $string.Replace("<stagename>", $obj.StageName)
	$string = $string.Replace("<place>", (getPlaceFull $obj.Place))
	$string = $string.Replace("<percent>", ($obj."Stage %").Replace(" ",""))
	$string = $string.Replace("<hitfactor>", $obj."Hit Factor")
	$string = $string.Replace("<div>", $($hash_division_friendly.Get_Item($shooter_division)))
	$string = $string.Replace("<classplace>", (getPlaceFull $obj.ClassPlace))
	$string = $string.Replace("<classpercent>", $obj.ClassPercent)
	$string = $string.Replace("<class>", $shooter_class)
	
	$string
}

#Define variables
$results_base_dir = "C:\temp"
$results_dir = "$results_base_dir\uspsa_results"
$filename = "test"
$base_path = "$results_base_dir\$filename"
$timeout_num = 5
$timeout_sleep_sec = 1
$match_name = "MRCS"
$date = Get-Date -format "yyyyMMdd-hhmmss"

# Keep track of what the practiscore website uses for division short names in their URLs
#$hash_division = @{
#"Production" = "prod"
#"Limited" = "ltd"
#"Open" = "open"
#"Limited10" = "ltdten"
#"Revolver" = "rev"
#"Single Stack" = "ss"
#}

$hash_division = @{
"PROD" = "prod"
"LTD" = "ltd"
"OPEN" = "open"
"L10" = "ltdten"
"REV" = "rev"
"SS" = "ss"
}

$hash_division_friendly = @{
"PROD" = "Production"
"LTD" = "Limited"
"OPEN" = "Open"
"L10" = "Limited10"
"REV" = "Revolver"
"SS" = "Single Stack"
}

$hash_division_extended = @{
"Production" = "production"
"Limited" = "limited10"
"Open" = "open"
"Limited10" = "limited"
"Revolver" = "revolver"
"Single Stack" = "singlestack"
}

$format = @"
Stage <stagenum> - <stagename>
<place> Place <div> - <percent>
<classplace> Place <div> <class> - <classpercent>
Hit Factor - <hitfactor>

"@


$match_name = Get-MatchName $URI
$overall_uri = Get-OverAllURI $URI
$shooter_name = Get-ShooterName $overall_uri
$shooter_division = Get-ShooterDivision $overall_uri
$shooter_class = Get-ShooterClass $overall_uri

$clean_match_name = $match_name.Replace(" ","")
$clean_shooter_name = $shooter_name.Replace(", ","")


$results_dir = "$results_dir\$clean_match_name"
$base_path = "$results_dir\$clean_match_name"
$output_path  = "$results_dir\$($match_name)_$($clean_shooter_name)_$date.txt"

# Create any temp or logging directory structures.
createFolder $results_base_dir
createFolder $results_dir

Write-Host "Getting match results for $Shooter"
Write-Host "Getting stage names for match"
$hash_stage_name = getStageName (getHTML $URI)
#Write-Host $hash_stage_name
$num_stages = $hash_stage_name.Count


if ($ClassCalc)
{
	$arr_individual_results = @()
	for ($i=1; $i -le $num_stages; $i++) {
		Write-Host "Getting results for stage $i."
		$stage_uri = getStageUri $URI $i $shooter_division
		$arr_results = getStageResults (getHTML $stage_uri) $i
		$arr_results | Export-CSV "$($base_path)-stage$($i).csv" -NoType
		$new_result = $arr_results | Where {$_.Name -eq $shooter_name}
		$new_result | Add-Member -MemberType NoteProperty -Name "StageName" -Value $hash_stage_name.Get_Item($i)
		$arr_individual_results += $new_result
	}

	$arr_individual_results | Export-CSV "$($base_path)-stagewhat.csv" -NoType

	Write-Host $output_path
	foreach ($obj in $arr_individual_results) {
		$place = getPlaceFull $obj.Place
		$stage_pct = ($obj."Stage %").Replace(" ","")
		$stage_name = $obj.StageName
		$stage_num = $obj.StageNum
		$hit_factor = $obj."Hit Factor"
		
$string = @"
Stage $stage_num - $stage_name
$place Place $($hash_division_friendly.Get_Item($shooter_division)) - $stage_pct
Hit Factor - $hit_factor

"@

		GenerateOutputWithFormat $obj $format | Out-File $output_path -Append
	}
}
else
{
	$arr_individual_results = @()
	for ($i=1; $i -le $num_stages; $i++) {
		Write-Host "Getting results for stage $i."
		$stage_uri = getStageUri $URI $i $shooter_division
		$arr_results = getStageResults (getHTML $stage_uri) $i
		$arr_results | Export-CSV "$($base_path)-stage$($i).csv" -NoType
		$arr_results = Convert-HitFactorToNum $arr_results
		$class_result = $arr_results | Where-Object {$_.Class -eq $shooter_class} | Sort "Hit Factor" -Descending
		
		
		$new_result = $class_result | Where {$_.Name -eq $shooter_name}
		
		$top_shooter = $class_result[0]
		[decimal]$top_hitfactor = [decimal]($top_shooter."Hit Factor")
		[decimal]$shooter_hitfactor = [decimal]($new_result."Hit Factor")
		$class_percent = [math]::Round(($shooter_hitfactor / $top_hitfactor * 100),2)
		
		[string]$class_place = Get-ClassPlace $class_result $shooter_name
		
		$new_result | Add-Member -MemberType NoteProperty -Name "StageName" -Value $hash_stage_name.Get_Item($i)
		$new_result | Add-Member -MemberType NoteProperty -Name "ClassPercent" -Value $class_percent
		$new_result | Add-Member -MemberType NoteProperty -Name "ClassPlace" -Value $class_place
		$arr_individual_results += $new_result
		$class_result | Export-CSV "$($base_path)-stage$($i)-class-$shooter_class.csv" -NoType
	}

	$arr_individual_results | Export-CSV "$($base_path)-individualresults.csv" -NoType

	Write-Host $output_path
	foreach ($obj in $arr_individual_results) {
		$class_place = getPlaceFull $obj.ClassPlace
		$class_percent = "$($obj.ClassPercent)%"
		$stage_name = $obj.StageName
		$stage_num = $obj.StageNum
		$hit_factor = $obj."Hit Factor"
		
$string = @"
Stage $stage_num - $stage_name
$class_place Place $($hash_division_friendly.Get_Item($shooter_division)) $shooter_class - $class_percent
Hit Factor - $hit_factor

"@

		GenerateOutputWithFormat $obj $format | Out-File $output_path -Append
	}


}






