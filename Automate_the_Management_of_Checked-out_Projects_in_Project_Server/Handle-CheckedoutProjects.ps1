<#

.SYNOPSIS
    Checks the Project Server database for checked-out projects and emails the user to check the Project back in, or forces the check-in.

.DESCRIPTION
    This script will query the Project Server draft database for the list of checked-out Projects.  If the checkout date exceeds the first defined value, the user who has 
    the Project checked-out will be sent an email asking them to check it in.  If the checkout date exceeds the second defined value, the Project will be checked-in by force
    and the user who had the Project checked out will be notified by email. 
    The variables in the Variables script region should be set as required.
    This script requires read permissions to the draft database as well as permission to check-in a Project plan.

.PARAMETER DaysUntilEmailUser
    The number of days the project plan can be checked out before the user is sent an email asking them to check it in.  The default is 7.

.PARAMETER DaysUntilForceCheckin
    The number of days the project plan can be checked out before the check-in is forced.  The default is 28.

.EXAMPLE
    .\Handle-CheckedoutProjects.ps1
    This will check the Project Server database for checked out projects, and if there are any that are checked out more than 28 days, they will be checked-in by force and
    the user notified.  If there are any that are checked out between 7 and 28 days, the user will be emailed advising them to check the Project back in.

.EXAMPLE
    .\Handle-CheckedoutProjects.ps1 -DaysUntilEmailUser 14 -DaysUntilForecCheckin 20
    This will check the Project Server database for checked out projects, and if there are any that are checked out more than 20 days, they will be checked-in by force and
    the user notified.  If there are any that are checked out between 14 and 20 days, the user will be emailed advising them to check the Project back in.

.NOTES
    Script name: Handle-CheckedoutProjects.ps1
    Author:      Trevor Jones
    Contact:     @trevor_smsagent
    DateCreated: 2015-02-17
    Link:        http://smsagent.wordpress.com

#>


[CmdletBinding(SupportsShouldProcess=$True)]
    param
        (
        [parameter(Mandatory=$False, HelpMessage="Number of days project is checked out before sending email")]
            [Int]$DaysUntilEmailUser = 7,
        [parameter(Mandatory=$False, HelpMessage="Number of days project is checked out before forcing check in")]
            [Int]$DaysUntilForceCheckin = 28
        )

# Script starts here

#region Variables

# Mail server info
$smtpserver = "mysmtpserver"
$From = "ProjectServer@contoso.com"
$admin = "Iam.admin@contoso.com"

# Database info
$dataSource = “sqlserver\instance”
$database = “PS_2010_PWA_DRAFT_90_DB”

# Project Server PWA URL
$ProjectServerURL = "http://project/PWA"

# Location of temp file for email message body (will be removed after)
$msgfile = "$env:TEMP\mailmessage.txt"

# Do not change
$ErrorActionPreference = "Stop"

#endregion



#region Functions

function New-Table (
$Topic1,
$Topic2,
$Topic3

)
{ 
       Add-Content $msgfile "<style>table {border-collapse: collapse;font-family: ""Trebuchet MS"", Arial, Helvetica, sans-serif;}"
       Add-Content $msgfile "h2 {font-family: ""Trebuchet MS"", Arial, Helvetica, sans-serif;}"
       Add-Content $msgfile "th, td {font-size: 1em;border: 1px solid #87ceeb;padding: 3px 7px 2px 7px;}"
       Add-Content $msgfile "th {font-size: 1.2em;text-align: left;padding-top: 5px;padding-bottom: 4px;background-color: #87ceeb;color: #ffffff;}</style>"
       Add-Content $msgfile "<p><table>"
       Add-Content $msgfile "<tr><th>$Topic1</th><th>$Topic2</th><th>$Topic3</th></tr>"
}

function New-TableRow (
$col1, 
$col2,
$col3

)
{
Add-Content $msgfile "<tr><td>$col1</td><td>$col2</td><td>$col3</td></tr>"
}

function New-TableEnd {
Add-Content $msgfile "</table></p>"}

#endregion



#region SQL

# Open a connection
Write-Verbose "Connecting to SQL Server $($dataSource) and database $($database)"
$connectionString = “Server=$dataSource;Database=$database;Integrated Security=SSPI;”
#$connectionString = "Server=$dataSource;Database=$database;uid=ProjServer_Read;pwd=Pa$$w0rd;Integrated Security=false"
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString
$connection.Open()

# Define the SQL query
$Query = "
SELECT PROJ_UID, PROJ_NAME, PROJ_PROP_AUTHOR, PROJ_CHECKOUTDATE, PROJ_SESSION_DESCRIPTION, PROJ_TYPE, res.RES_NAME, res.WRES_EMAIL 
FROM dbo.MSP_projects proj
inner join dbo.MSP_RESOURCES_PUBLISHED_VIEW res on proj.PROJ_CHECKOUTBY = res.RES_UID
WHERE PROJ_CHECKOUTBY is not null
and DATEDIFF(day,PROJ_CHECKOUTDATE,GETDATE()) > $DaysUntilEmailUser 
ORDER BY PROJ_CHECKOUTDATE Desc
"
# Run the query
Write-Verbose "Executing SQL query"
$command = $connection.CreateCommand()
$command.CommandText = $query
$reader = $command.ExecuteReader()

# Put results into a table
Write-Verbose "Adding results to a table"
$table = new-object “System.Data.DataTable”
$table.Load($reader)
 
# Close the connection
Write-Verbose "Closing SQL connection"
$connection.Close()

#endregion



#region SendEmails

if ($table.rows.count -gt 0)
    {
        foreach ($row in $table.rows)
            {
                # Create file
                New-Item $msgfile -ItemType file -Force | Out-Null

                # If checkoutdate has been exceeded, force checkin and email user
                if ($row.PROJ_CHECKOUTDATE -lt (get-date).AddDays(-$DaysUntilForceCheckin))
                    {
                        try
                            {
                                Write-verbose "Checking in Project: $($row.PROJ_NAME)"
                                $svcPSProxy = New-WebServiceProxy -uri "$ProjectServerURL/_vti_bin/PSI/Project.asmx?wsdl" -useDefaultCredential
                                $projId = [System.Guid]$row.PROJ_UID.Guid
                                $svcPSProxy.QueueCheckInProject([System.Guid]::NewGuid() , $projId, "true",[System.Guid]::NewGuid(),"Big Brother")
                            }
                        catch {continue}
                        
                        # Add html header
                        Add-Content $msgfile "<style>h3 {font-family: ""Trebuchet MS"", Arial, Helvetica, sans-serif;}</style>"
                        Add-Content $msgfile "<h3>Dear $($row.RES_NAME.Split(" ")[0]),</h3>"
                        Add-Content $msgfile "<h3>The following project plan has been checked out to you since $((Get-date $($row.PROJ_CHECKOUTDATE) -Format "dd MMMM yyyy").ToString()).  The check-in has now been forced.</h3>"
                        Add-Content $msgfile "<p></p>"
                        
                        # Create a new html table
                        New-Table -Topic1 "Project Name" -Topic2 "Project Author" -Topic3 "Session"
            
                        # Populate table rows with project plan details
                        $Proj_name = $row.PROJ_NAME | Out-String
                        $Proj_Auth = $row.PROJ_PROP_AUTHOR | Out-String
                        $Proj_Sessions = $row.PROJ_SESSION_DESCRIPTION | Out-String

                        New-TableRow -col1 $Proj_name -col2 $Proj_Auth -col3 $Proj_Sessions
        
                        # Add html table to file
                        New-TableEnd
                
                        # Set email body content
                        $body = Get-Content $msgfile

                        # Email user
                        Write-Verbose "Sending ""Check-in forced"" email to $($row.WRES_EMAIL) for Project $($row.PROJ_NAME) checked out on $($row.PROJ_CHECKOUTDATE)"
                        Send-MailMessage -To $row.WRES_EMAIL -Bcc $admin -Subject "PROJECT CHECK-IN: $($Row.PROJ_NAME)" -Body "$body" -From $from -SmtpServer $smtpserver -BodyAsHtml

                        # Delete tempfile 
                        Remove-Item $msgfile
                    }
                
                # If checkoutdate is less than forcecheckin date, email user
                if ($row.PROJ_CHECKOUTDATE -ge (get-date).AddDays(-$DaysUntilForceCheckin) -and $row.PROJ_CHECKOUTDATE -lt (get-date).AddDays(-$DaysUntilEmailUser))
                    {
                            
                        # Add html header
                        Add-Content $msgfile "<style>h3 {font-family: ""Trebuchet MS"", Arial, Helvetica, sans-serif;}</style>"
                        Add-Content $msgfile "<h3>Dear $($row.RES_NAME.Split(" ")[0]),</h3>"
                        Add-Content $msgfile "<h3>The following project plan has been checked out to you since $((Get-date $($row.PROJ_CHECKOUTDATE) -Format "dd MMMM yyyy").ToString()).  Please check in the project ASAP.</h3>"
                        Add-Content $msgfile "<p></p>"
                         
                        # Create a new html table
                        New-Table -Topic1 "Project Name" -Topic2 "Project Author" -Topic3 "Session" 
                                                
                        # Populate table rows with project plan details
                        $Proj_name = $row.PROJ_NAME | Out-String
                        $Proj_Auth = $row.PROJ_PROP_AUTHOR | Out-String
                        $Proj_Sessions = $row.PROJ_SESSION_DESCRIPTION | Out-String
        
                        New-TableRow -col1 $Proj_name -col2 $Proj_Auth -col3 $Proj_Sessions
        
                        # Add html table to file
                        New-TableEnd
                
                        # Set email body content
                        $body = Get-Content $msgfile

                        # Email user 
                        Write-Verbose "Sending ""Check-in requested"" email to $($row.WRES_EMAIL) for Project $($row.PROJ_NAME) checked out on $($row.PROJ_CHECKOUTDATE)"           
                        Send-MailMessage -To $row.WRES_EMAIL -Bcc $admin -Subject "OVERDUE PROJECT CHECK-OUT: $($Row.PROJ_NAME)" -Body "$body" -From $from -SmtpServer $smtpserver -BodyAsHtml
        
                        # Delete tempfile 
                        Remove-Item $msgfile
                    }
            }
    }

#endregion