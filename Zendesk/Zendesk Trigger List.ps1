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
	$zendeskurl="https://$zendeskprefix.zendesk.com/api/v2/triggers.json?page=$currentpage"
	$triggerlist=$w.downloadstring($zendeskurl) | ConvertFrom-Json
	foreach ($trigger in $triggerlist.triggers){$al.add($trigger)}
	$currentpage++
} while ($triggerlist.next_page -ne $null)