param($installPath, $toolsPath, $package, $project)

Import-Module (Join-Path $toolsPath common.psm1) -Force

try {

    # Indicates if current project is a VB project
    $IsVbProject = ($project.CodeModel.Language -eq [EnvDTE.CodeModelLanguageConstants]::vsCMLanguageVB)
    
    if ($IsVbProject) {

        # Remove added Imports for VB project
        $VbImports | ForEach-Object {
            if ($project.Object.Imports -contains $_) {
                $project.Object.Imports.Remove($_)
            }
        }
    }
    
    # Remove references
    Foreach ($spRef in $ReferencesToRemoveWhenUninstall) {
        $project.Object.References | Where-Object { $_.Identity -eq $spRef } | ForEach-Object { $_.Remove() }
    }
    
} catch {

    Write-Host "Error uninstalling package: " + $_.Exception -ForegroundColor Red
    exit
}
# SIG # Begin signature block
# MIIh9wYJKoZIhvcNAQcCoIIh6DCCIeQCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAN/6dBHJ45hY+Y
# N8R45TKbxt2wbCsKlOEE8Hxkc3JCuqCCC4MwggULMIID86ADAgECAhMzAAAAM1b2
# lB2ajL3lAAAAAAAzMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTAwHhcNMTMwOTI0MTczNTU1WhcNMTQxMjI0MTczNTU1WjCBgzEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjENMAsGA1UECxMETU9Q
# UjEeMBwGA1UEAxMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMIIBIjANBgkqhkiG9w0B
# AQEFAAOCAQ8AMIIBCgKCAQEAs9KaOIfw6Oly8PBcJp2mW2pAcbiYWLBfGneq+Oed
# i8Vc8IrjSTO4bEGak9UTxlyKNykoTjwpF275u22O3FPFEQPJU96Y8PFN7E2x8gh4
# 6ftxxmL9XCqnZGd4YJ+qhW3OPuJq9DLc14DJiKAxmHE69CH3N65QJId20RHix/47
# PaEYkBalXwSZ6JLjG9MJSFwmBVUb3WilzUsPv/XM3lWltHUqcbSZwjsM5NKR2HKK
# +eyHIqxqWb90NUky2K0jSbVnEJgQy9TIljp84OA+7ei+v2Lo4dJ7eAYGodazlE1W
# BQ2vCD7ItSKc/m0QL+tjGxW5kCeRZ/sSHyvcdveB1CphyQIDAQABo4IBejCCAXYw
# HwYDVR0lBBgwFgYIKwYBBQUHAwMGCisGAQQBgjc9BgEwHQYDVR0OBBYEFPBHESyD
# Hm5wg0qUmlqkIi/UPOxLMFEGA1UdEQRKMEikRjBEMQ0wCwYDVQQLEwRNT1BSMTMw
# MQYDVQQFEyozODA3NisxMzVlOTk3ZC0yZmUyLTQ3MWMtYjIxYy0wY2VmNjA1OGU5
# ZjYwHwYDVR0jBBgwFoAU5vxfe7siAFjkck619CF0IzLm76wwVgYDVR0fBE8wTTBL
# oEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMv
# TWljQ29kU2lnUENBXzIwMTAtMDctMDYuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggr
# BgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWND
# b2RTaWdQQ0FfMjAxMC0wNy0wNi5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0B
# AQsFAAOCAQEAUCzVYWVAmy0CuJ1srWZf0GzTE7bv6EBw3KVMIUi+aQDV1Cmyip6P
# 0aaVqwn2IU4fZCm9cISyrZvvZtsBgZo427YflDWZwXnJVdOhfnUfXD0Ql0G3/eXq
# nwZrQED6XhbKSWXC6g3R47bWLMO2FxrD+oC81yC5iYGvJFCy+iWW7T7Sp2MMr8nZ
# XUmh7VwqxLmESRL9SG0I1jBJeiw3np61RvhG9K7I3ADQAlAwgs07dOphCztGdya7
# LMU0aPEHo4nShwMWGGISjVayRZ3K3KlQQgWDzrgF4alEgf5eHQObN3ZA01YoN2Ir
# J5IcVCEDiAcMbEMVqFPt6srBJveymDXpPDCCBnAwggRYoAMCAQICCmEMUkwAAAAA
# AAMwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1
# dGhvcml0eSAyMDEwMB4XDTEwMDcwNjIwNDAxN1oXDTI1MDcwNjIwNTAxN1owfjEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWlj
# cm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMDCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBAOkOZFB5Z7XE4/0JAEyelKz3VmjqRNjPxVhPqaV2fG1FutM5
# krSkHvn5ZYLkF9KP/UScCOhlk84sVYS/fQjjLiuoQSsYt6JLbklMaxUH3tHSwoke
# cZTNtX9LtK8I2MyI1msXlDqTziY/7Ob+NJhX1R1dSfayKi7VhbtZP/iQtCuDdMor
# sztG4/BGScEXZlTJHL0dxFViV3L4Z7klIDTeXaallV6rKIDN1bKe5QO1Y9OyFMjB
# yIomCll/B+z/Du2AEjVMEqa+Ulv1ptrgiwtId9aFR9UQucboqu6Lai0FXGDGtCpb
# nCMcX0XjGhQebzfLGTOAaolNo2pmY3iT1TDPlR8CAwEAAaOCAeMwggHfMBAGCSsG
# AQQBgjcVAQQDAgEAMB0GA1UdDgQWBBTm/F97uyIAWORyTrX0IXQjMubvrDAZBgkr
# BgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUw
# AwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBN
# MEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0
# cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoG
# CCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01p
# Y1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDCBnQYDVR0gBIGVMIGSMIGPBgkrBgEE
# AYI3LgMwgYEwPQYIKwYBBQUHAgEWMWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9Q
# S0kvZG9jcy9DUFMvZGVmYXVsdC5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcA
# YQBsAF8AUABvAGwAaQBjAHkAXwBTAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZI
# hvcNAQELBQADggIBABp071dPKXvEFoV4uFDTIvwJnayCl/g0/yosl5US5eS/z7+T
# yOM0qduBuNweAL7SNW+v5X95lXflAtTx69jNTh4bYaLCWiMa8IyoYlFFZwjjPzwe
# k/gwhRfIOUCm1w6zISnlpaFpjCKTzHSY56FHQ/JTrMAPMGl//tIlIG1vYdPfB9XZ
# cgAsaYZ2PVHbpjlIyTdhbQfdUxnLp9Zhwr/ig6sP4GubldZ9KFGwiUpRpJpsyLcf
# ShoOaanX3MF+0Ulwqratu3JHYxf6ptaipobsqBBEm2O2smmJBsdGhnoYP+jFHSHV
# e/kCIy3FQcu/HUzIFu+xnH/8IktJim4V46Z/dlvRU3mRhZ3V0ts9czXzPK5UslJH
# asCqE5XSjhHamWdeMoz7N4XR3HWFnIfGWleFwr/dDY+Mmy3rtO7PJ9O1Xmn6pBYE
# AackZ3PPTU+23gVWl3r36VJN9HcFT4XG2Avxju1CCdENduMjVngiJja+yrGMbqod
# 5IXaRzNij6TJkTNfcR5Ar5hlySLoQiElihwtYNk3iUGJKhYP12E8lGhgUu/WR5mg
# gEDuFYF3PpzgUxgaUB04lZseZjMTJzkXeIc2zk7DX7L1PUdTtuDl2wthPSrXkizO
# N1o+QEIxpB8QCMJWnL8kXVECnWp50hfT2sGUjgd7JXFEqwZq5tTG3yOalnXFMYIV
# yjCCFcYCAQEwgZUwfjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEoMCYGA1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMAITMwAA
# ADNW9pQdmoy95QAAAAAAMzANBglghkgBZQMEAgEFAKCBuDAZBgkqhkiG9w0BCQMx
# DAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkq
# hkiG9w0BCQQxIgQgYpOtBb7+0rbI0S3pbSEkF8U8URvtfTZfOz5+KgQnVQ4wTAYK
# KwYBBAGCNwIBDDE+MDygIoAgAHUAbgBpAG4AcwB0AGEAbABsAF8AMQA0AC4AcABz
# ADGhFoAUaHR0cDovL21pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEBBQAEggEAHCuR
# k0stlDGx2jLTIWO5GaGw4/T4U86FHUPzgaypcimm+D/9mkJzW8nsAbWSTs3p8jcd
# 3XmDWXz3BBEIvR9TTvGMS2GrLQxys9t79hEJ5KPFixS/l5y3EvxbSmuCkoK/f8Nf
# zrQIHgY6awwYqDFBQ9qWqCxXacM761hQfK52xsI3Lejt2GmOt/9PaI+sgUZscUFA
# ROHTOhp+uzJ9U2g+Waf37UMSm+aXDPZIoUMc2Cl9F52NLMEj0Cs8CCsvJ0CxSJiT
# 7NbXOhpN0fyPp59LEr6KbQDaijk2o+1bulW0PcQcYi0PXYsDEHySoDiynH8SCBaQ
# si78Vq2DZ443vYdo4qGCE0owghNGBgorBgEEAYI3AwMBMYITNjCCEzIGCSqGSIb3
# DQEHAqCCEyMwghMfAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggE9BgsqhkiG9w0BCRAB
# BKCCASwEggEoMIIBJAIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCDH
# RkQbO3pQP/EtlTI3ay1/PBimxW86yWv3WhQualbLzgIGU0/2wxJQGBMyMDE0MDQy
# MjA1NTIyMy40MjNaMAcCAQGAAgH0oIG5pIG2MIGzMQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMScwJQYDVQQLEx5uQ2lw
# aGVyIERTRSBFU046QzBGNC0zMDg2LURFRjgxJTAjBgNVBAMTHE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFNlcnZpY2Wggg7NMIIGcTCCBFmgAwIBAgIKYQmBKgAAAAAAAjAN
# BgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
# b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
# dGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9y
# aXR5IDIwMTAwHhcNMTAwNzAxMjEzNjU1WhcNMjUwNzAxMjE0NjU1WjB8MQswCQYD
# VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEe
# MBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3Nv
# ZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCC
# AQoCggEBAKkdDbx3EYo6IOz8E5f1+n9plGt0VBDVpQoAgoX77XxoSyxfxcPlYcJ2
# tz5mK1vwFVMnBDEfQRsalR3OCROOfGEwWbEwRA/xYIiEVEMM1024OAizQt2TrNZz
# MFcmgqNFDdDq9UeBzb8kYDJYYEbyWEeGMoQedGFnkV+BVLHPk0ySwcSmXdFhE24o
# xhr5hoC732H8RsEnHSRnEnIaIYqvS2SJUGKxXf13Hz3wV3WsvYpCTUBR0Q+cBj5n
# f/VmwAOWRH7v0Ev9buWayrGo8noqCjHw2k4GkbaICDXoeByw6ZnNPOcvRLqn9Nxk
# vaQBwSAJk3jN/LzAyURdXhacAQVPIk0CAwEAAaOCAeYwggHiMBAGCSsGAQQBgjcV
# AQQDAgEAMB0GA1UdDgQWBBTVYzpcijGQ80N7fEYbxTNoWoVtVTAZBgkrBgEEAYI3
# FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAf
# BgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBH
# hkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNS
# b29DZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUF
# BzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0Nl
# ckF1dF8yMDEwLTA2LTIzLmNydDCBoAYDVR0gAQH/BIGVMIGSMIGPBgkrBgEEAYI3
# LgMwgYEwPQYIKwYBBQUHAgEWMWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9QS0kv
# ZG9jcy9DUFMvZGVmYXVsdC5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBs
# AF8AUABvAGwAaQBjAHkAXwBTAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcN
# AQELBQADggIBAAfmiFEN4sbgmD+BcQM9naOhIW+z66bM9TG+zwXiqf76V20ZMLPC
# xWbJat/15/B4vceoniXj+bzta1RXCCtRgkQS+7lTjMz0YBKKdsxAQEGb3FwX/1z5
# Xhc1mCRWS3TvQhDIr79/xn/yN31aPxzymXlKkVIArzgPF/UveYFl2am1a+THzvbK
# egBvSzBEJCI8z+0DpZaPWSm8tv0E4XCfMkon/VWvL/625Y4zu2JfmttXQOnxzplm
# kIz/amJ/3cVKC5Em4jnsGUpxY517IW3DnKOiPPp/fZZqkHimbdLhnPkd/DjYlPTG
# pQqWhqS9nhquBEKDuLWAmyI4ILUl5WTs9/S/fmNZJQ96LjlXdqJxqgaKD4kWumGn
# Ecua2A5HmoDF0M2n0O99g/DhO3EJ3110mCIIYdqwUB5vvfHhAN/nMQekkzr3ZUd4
# 6PioSKv33nJ+YWtvd6mBy6cJrDm77MbL2IK0cs0d9LiFAR6A+xuJKlQ5slvayA1V
# mXqHczsI5pgt6o3gMy4SKfXAL1QnIffIrE7aKLixqduWsqdCosnPGUFN4Ib5Kpqj
# EWYw07t0MkvfY3v1mYovG8chr1m1rtxEPJdQcdeh0sVV42neV8HR3jDA/czmTfsN
# v11P6Z0eGTgvvM9YBS7vDaBQNdrvCScc1bN+NR4Iuto229Nfj950iEkSMIIE2jCC
# A8KgAwIBAgITMwAAACiQZ7kEsDxuZgAAAAAAKDANBgkqhkiG9w0BAQsFADB8MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0xMzAzMjcyMDEzMTNaFw0xNDA2
# MjcyMDEzMTNaMIGzMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MQ0wCwYDVQQLEwRNT1BSMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046QzBGNC0z
# MDg2LURFRjgxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Uw
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDdpUi/akidSiGckmve4C3c
# 5GP4zLmJxMcbvee10/vtrs8x/vNmsEQD2plnCFq/dQYiEYnQZ1LM+s+SN0Xo+vG9
# M9PMc+O4IaSgFX3LL8QDBdo/lnPTWeWYTQtWhi+dR9HWX52R6ceE2ZVrMky0awBS
# 4EHTPGl0qM7MfWidUlXmcH8UB6KeZ7CGRPMzP3Ndxij4F19SAS1EL9bteAi45Tsv
# wLnDS8O3Oy/TprWcsUhK3TIJVqEbS1rTqiYnDBJDYMVq19pADWCYiUG7k3Pdv/7E
# jFvO+lUnyk1Nmm99EWyxRyOwTHxsfwahdIIfUngY6QYaFlCawzrdgYH3mydyIX91
# AgMBAAGjggEbMIIBFzAdBgNVHQ4EFgQU3JgInXnRBLKLR8Nx0Izns+awU50wHwYD
# VR0jBBgwFoAU1WM6XIoxkPNDe3xGG8UzaFqFbVUwVgYDVR0fBE8wTTBLoEmgR4ZF
# aHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljVGlt
# U3RhUENBXzIwMTAtMDctMDEuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcw
# AoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNUaW1TdGFQ
# Q0FfMjAxMC0wNy0wMS5jcnQwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEF
# BQcDCDANBgkqhkiG9w0BAQsFAAOCAQEAgiLztz1kfhJL/Cb84OS30MQUTgn+q1aa
# 0VqYpr6MQR6UtDK+hLS3RXbj72AYJIeoz+m00VQpvMrkyxJ7wPHUDp8xMxsRP3o7
# 3d0CqhjKyjz6luNsu6+7yYQ+x9gMhctyCwEbpPUxERAMRaVaSJl+2r5Fhte6TeSB
# /9NYCnZlBlkv9sJCzwTJqxv6YZ3185hJcLFJ0GTEIejuYBdTfusC2miVi/UKPAHb
# o7WYFFF0nlPp2nKYZqBfKc+Prx+CnNPr5vFMG1T46DLcwRXDrCpudAUWg+NEmJ/L
# 7+gweX+vUqU6H99lx43+J9hHGZIItIs0jmknNxoC9pGzlSL/CEgq/qGCA3YwggJe
# AgEBMIHjoYG5pIG2MIGzMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMQ0wCwYDVQQLEwRNT1BSMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046QzBG
# NC0zMDg2LURFRjgxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZp
# Y2WiJQoBATAJBgUrDgMCGgUAAxUA8120HsdfO2ZOZQ7emART9hWnH0SggcIwgb+k
# gbwwgbkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xDTALBgNV
# BAsTBE1PUFIxJzAlBgNVBAsTHm5DaXBoZXIgTlRTIEVTTjpCMDI3LUM2RjgtMUQ4
# ODErMCkGA1UEAxMiTWljcm9zb2Z0IFRpbWUgU291cmNlIE1hc3RlciBDbG9jazAN
# BgkqhkiG9w0BAQUFAAIFANcAKbowIhgPMjAxNDA0MjEyMzM1MjJaGA8yMDE0MDQy
# MjIzMzUyMlowdDA6BgorBgEEAYRZCgQBMSwwKjAKAgUA1wApugIBADAHAgEAAgIK
# 0jAHAgEAAgIXVzAKAgUA1wF7OgIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgorBgEE
# AYRZCgMBoAowCAIBAAIDFuNgoQowCAIBAAIDB6EgMA0GCSqGSIb3DQEBBQUAA4IB
# AQBh0uH0DLnszayqwL5/85JdIlh338UL6EhehCRJcF8q6IDkP4mrtso1tXFgnPqz
# GwjE5QZ/wkrKTh8oIsJb2rAG5zmDdHBp3ph4WCB+41YR8YKn1RUFvwyLwRDzGkNg
# n6KouyL72Sq40pLIT1Winvo9o8RBolAbaIuJ1s7K6U45K6mU7n9pS0jdvbiV8odt
# 5qAvDoYnKg3mmoOtAGx3NEVHOAek4PhtxNIOMqtKEa/P/feraexFgRygMLudWkx0
# oV//4NFzM+TeGJZ82YPl7e8uhV5ya/PZlrewXN/hqWA6lrQj2KXYKOe3YuUAyWFF
# bfXycc2T5ijt157R33kHOrYNMYIC9TCCAvECAQEwgZMwfDELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBIDIwMTACEzMAAAAokGe5BLA8bmYAAAAAACgwDQYJYIZIAWUDBAIB
# BQCgggEyMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQx
# IgQgS8EIBsYPTlJ2dHCm6AJ0vBK7tg+mpTms350XE9cWXP4wgeIGCyqGSIb3DQEJ
# EAIMMYHSMIHPMIHMMIGxBBTzXbQex187Zk5lDt6YBFP2FacfRDCBmDCBgKR+MHwx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAAKJBnuQSwPG5mAAAAAAAo
# MBYEFKX5BwDMaYek4P61NlbUIACDx0JiMA0GCSqGSIb3DQEBCwUABIIBAIWHgysy
# XWXuwjbq+cyOuMGrDUMTzzGj7nZm37LWSnFBbIhRN954BYrdqVgHnv/RBUQNOJop
# 3N+klVeUUmMKYb/XN8QNfhUMTyXiD0V4+V+1H1uxpWJEM4IN0TBZgEi29HYBHz0G
# b8eGq2PahSnufU9mogqz+HoiWcDlvDMwtD8RGWKsyeaxMXcD/aOS3T8v+VBeBCYP
# 2IYVeb+NkaCBagBLUS07CEoy2H5vVcnhatRCN5XxNN3AFUv296/+s43ujAKlWBDh
# XaJmqB9pspm9LdW2h6zDMM3uHXGzYttvaFMejALAFm7lcg/Vi8k0S5fFGn4YsqXD
# /0JlCjfS8Pyfpms=
# SIG # End signature block
