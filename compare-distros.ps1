# Custom functions that will add or remove users based on supplied information below
# These functions will scroll through each individual member of the supplied arrays and add/remove accordingly
#
function addMembers($distro, $users) {
    foreach ($member in $users) {
        Add-DistributionGroupMember -Identity $distro -Member $member
    }
}

function removeMembers($distro, $users) {
    foreach ($member in $users) {
        Remove-DistributionGroupMember -Identity $distro -Member $member
    }
}

# Distribution Lists that need to be changed, provided by reporter
#
$directorDistro = "{DISTRO_NAME1}"
$seniorDistro = "{DISTRO_NAME2}"
$managingDistro = "{DISTRO_NAME3}"

# Grabs the email addresses for each user that is currently in the distribution lists and stores them as an array
# 
$currentDirector = @((Get-DistributionGroupMember -Identity $directorDistro).PrimarySmtpAddress)
$currentSenior = @((Get-DistributionGroupMember -Identity $seniorDistro).PrimarySmtpAddress)
$currentManaging = @((Get-DistributionGroupMember -Identity $managingDistro).PrimarySmtpAddress)

# Updated lists, provided by reporter
# Make sure to store the lists as arrays, for the custom function
#
$newDirector = @({USER_LIST1})
$newSenior = @({USER_LIST2})
$newManaging = @({USER_LIST3})

# Arrays that will contain the differences between the new and current lists
# The add{NAME} variables will store every value of $new{NAME} that is not contained in the $current{NAME} list
# The remove{NAME} variables will store every value of $current{NAME} that is not contained in the $new{NAME} list
#
$addDirectors = @($newDirector | Where-Object {$currentDirector -NotContains $_})
$removeDirectors = @($currentDirector | Where-Object {$newDirector -NotContains $_})

$addSeniors = @($newSenior | Where-Object {$currentSenior -NotContains $_})
$removeSeniors = @($currentSenior | Where-Object {$newSenior -NotContains $_})

$addManaging = @($newManaging | Where-Object {$currentManaging -NotContains $_})
$removeManaging = @($currentManaging | Where-Object {$newManaging -NotContains $_})

# Call custom functions to add and remove specific users
# When calling a function with more than 1 parameter, do not use parenthesis or commas
# WRONG: foo($var1, $var2)
# CORRECT: foo $var1 $var2
#
addMembers $directorDistro $addDirectors
removeMembers $directorDistro $removeDirectors

addMembers $seniorDistro $addSeniors
removeMembers $seniorDistro $removeSeniors

addMembers $managingDistro $addManaging
removeMembers $managingDistro $removeManaging
