$w = new-object net.webclient
$al = New-Object System.Collections.ArrayList

$email = "username@example.org"
$username = "$email/token"
$password = "base64APITOKEN"
$zendeskprefix= "subdomain"

$w.Credentials = new-object System.Net.NetworkCredential($username, $password)

$currentpage = 1

do {
	#increment pages as long as you have pages that have been fully processed.
	$zendeskurl="https://$zendeskprefix.zendesk.com/api/v2/users.json?page=$currentpage"
	$userlist=$w.downloadstring($zendeskurl) | ConvertFrom-Json
	foreach ($user in $userlist.users){$al.add($user)}
	$currentpage++
} while ($userlist.next_page -ne $null)