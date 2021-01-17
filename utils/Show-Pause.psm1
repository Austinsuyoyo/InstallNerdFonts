<#
.Synopsis
Displays a visual representation of a calendar.

.Description
Displays a visual representation of a calendar. This function supports multiple months
and lets you highlight specific date ranges or days.

.Parameter Start
The first month to display.

.Parameter End
The last month to display.

.Parameter FirstDayOfWeek
The day of the month on which the week begins.

.Parameter HighlightDay
Specific days (numbered) to highlight. Used for date ranges like (25..31).
Date ranges are specified by the Windows PowerShell range syntax. These dates are
enclosed in square brackets.

.Parameter HighlightDate
Specific days (named) to highlight. These dates are surrounded by asterisks.

.Example
# Show a default display of this month.
Show-Calendar

.Example
# Display a date range.
Show-Calendar -Start "March, 2010" -End "May, 2010"

.Example
# Highlight a range of days.
Show-Calendar -HighlightDay (1..10 + 22) -HighlightDate "December 25, 2008"
#>
function Show-Pause {
    param(
        [Parameter()]
        [string]
        $Message
    )
    # Check if running Powershell ISE
    if ($psISE) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("$Message")
    }
    else {
        Write-Host "$Message" -ForegroundColor Red
        [void]($host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown"))
    }
}
Export-ModuleMember -Function Show-Pause