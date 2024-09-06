# Prompt for the user's email address
$UserEmail = Read-Host -Prompt "Enter the user's email address"

# Check if the email address is an active account
$User = Get-Mailbox -Identity $UserEmail -ErrorAction SilentlyContinue

if ($User -eq $null) {
    Write-Output "The email address '$UserEmail' is not an active account."
} else {
    # Get all shared mailboxes the user has access to
    $SharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Get-MailboxPermission -User $UserEmail

    # Create a custom object to store the results
    $Results = @()

    foreach ($Mailbox in $SharedMailboxes) {
        # Get the full email address of the shared mailbox
        $MailboxEmail = (Get-Mailbox $Mailbox.Identity).PrimarySmtpAddress

        # Get the permissions for the shared mailbox
        $Permissions = Get-MailboxPermission -Identity $Mailbox.Identity

        # Filter out the default permissions and count the unique users
        $MemberCount = ($Permissions | Where-Object { $_.User -notlike "NT AUTHORITY\SELF" -and $_.User -notlike "NT AUTHORITY\SYSTEM" -and $_.User -notlike "S-1-5-*" }).User | Sort-Object -Unique | Measure-Object | Select-Object -ExpandProperty Count

        # Add the result to the custom object
        $Results += [PSCustomObject]@{
            Identity    = $MailboxEmail
            User        = $Mailbox.User
            AccessRights = $Mailbox.AccessRights
            MemberCount = $MemberCount
        }
    }

    # Output the results
    $Results | Select-Object Identity, User, AccessRights, MemberCount
}
