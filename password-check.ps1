get-aduser james.cadd -Properties samaccountname, passwordlastset, "msDS-UserPasswordExpiryTimeComputed" | Select-Object samaccountname, PasswordLastSet, @{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}