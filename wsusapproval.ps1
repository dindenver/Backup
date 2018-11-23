param
(
[Parameter(Position = 0, Mandatory = $true, HelpMessage="Approve patches downloaded after which date")]
[String]
$ApprovalGroup,
[DateTime]
$DateThreshold = $((get-date).adddays(-90))
)

Get-WsusUpdate -Approval Unapproved | where ({$_.datetime -gt $datethreshold})