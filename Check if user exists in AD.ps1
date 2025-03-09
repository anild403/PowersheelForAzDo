function Test-ADUser {
  param(
    [Parameter(Mandatory)]
    [String]
    $sAMAccountName
  )
  $null -ne ([ADSISearcher] "(sAMAccountName=$sAMAccountName)").FindOne()
}

([ADSISearcher] "(sAMAccountName=userid)").FindOne()