# Manager's name should be called when run the script as "firstname.lastname", how it's shown in AD
# Example: ./create-dynamic-distro.ps1 "james.cadd"
param($manager)

# This will take the manager's name and capitilize it and remove the period, for the direct name of the dynamic distro
# Example: james.cadd -> James.Cadd -> James Cadd
$managerName = ((Get-Culture).TextInfo.ToTitleCase($manager)).replace("."," ")

# This will take the manager's name and remove the period, for the email address format that we have in place already
# Example: james.cadd -> jamescadd
$emailManagerName = $manager.replace(".","")

# This will create the name of the dynamic distro
# Example: "James Cadd Direct Reports"
$DistributionName = $managerName + " Direct Reports"

# This will create the primary email address
# Example: jamescadd-directs@caddnation.com
$DistributionEmail = $emailManagerName + "-directs@caddnation.com"

# You must be authenticated by using the Connect-ExchangeOnline command to continue

# This will grab the distinguished name from the Exchange Server
# Example: "CN=James Cadd,OU=CaddNation.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=NAMPR01A008,DC=PROD,DC=OUTLOOK,DC=COM"
$DistinguishedName = (Get-Recipient -Identity $manager).DistinguishedName

# This will create the filter for the dynamic distro
# Example: "Manager -eq 'CN=James Cadd,OU=CaddNation.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=NAMPR01A008,DC=PROD,DC=OUTLOOK,DC=COM'"
$Filter = "Manager -eq '$DistinguishedName'"

# This will take everything that you have collected and create the dynamic distro
# Example: New-DynamicDistributionGroup -Name "James Cadd Direct Reports" -PrimarySmtpAddress "jamescadd-directs@caddnation.com" -RecipientFilter "Manager -eq 'CN=James Cadd,OU=CaddNation.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=NAMPR01A008,DC=PROD,DC=OUTLOOK,DC=COM'"
New-DynamicDistributionGroup -Name $DistributionName -PrimarySmtpAddress $DistributionEmail -RecipientFilter $Filter
