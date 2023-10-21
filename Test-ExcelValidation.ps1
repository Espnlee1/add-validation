#Import-Excel
$coachData = Import-Excel "nba.hof.xlsx"
 
#Create files for coaches based on column C (Coach) where coach is not like $null
foreach ($coach in $coachData.Where({$_."coach" -notlike $null}) | group "coach"){
 
    #Create output file name for each and thier players data.
    $newXlfile = "$($coach.name) HOF Palyers.xlsx"
 
    #rm short for remove item
    rm $newXlfile -ErrorAction SilentlyContinue
   
    #Export data for each coach into seperate Excel files
    $coach.group | Export-Excel $newXlfile -WorksheetName $coach.Name -FreezeTopRow -TableStyle Light8 -BoldTopRow -AutoFilter -AutoSize
 
    #Opening Excel Package to get the actual instace of the first sheet in $newXlfile workbooks and assign to the $ws variable
    $xl = Open-ExcelPackage $newXlfile
    $ws = $xl.Workbook.Worksheets[1]
 
    #Splatting for validation based on example
    $ValidationParams = @{
    Worksheet        = $ws
    ShowErrorMessage = $true
    ErrorStyle       = 'stop'
    ErrorTitle       = 'Invalid Data'
    }#end $ValidationParams
 
 
    #Additional Splatting for validation based on example
    $MoreValidationParams = @{
    Range          = 'D2:D100001'
    ValidationType = 'List'
    ValueSet       = @('YES', 'NO', 'MAYBE')
    ErrorBody      = "You must select an item from the list."
    }# End $MoreValidationParams
 
    #Apply data validation to Each file
    Add-ExcelDataValidationRule @ValidationParams @MoreValidationParams
 
 
} #End foreach
