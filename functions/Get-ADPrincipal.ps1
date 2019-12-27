function Get-ADPrincipal
{
#https://twitter.com/IISResetMe/status/1087302392364912645
  param([Security.Principal.SecurityIdentifier]$SID)

  $bytes = [byte[]]::new($SID.BinaryLength)
  $SID.GetBinaryForm($bytes,0)

  Get-ADObject -LDAPFilter "(&(objectSid=$($bytes.ForEach({'\{0:X2}'-f$_})-join'')))"
}