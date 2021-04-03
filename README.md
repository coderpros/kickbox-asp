<img src="https://coderpro.net/images/logos/coderPro_logo_rounded_extra-90x90.webp" align="right" />

[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![MIT License][license-shield]][license-url]
[![LinkedIn][linkedin-shield]][linkedin-url]
[![Twitter](https://img.shields.io/twitter/url/https/twitter.com/cloudposse.svg?style=social&label=Follow%20%40coderProNet)](https://twitter.com/coderProNet)
[![GitHub](https://img.shields.io/github/followers/coderpros?label=Follow&style=social)](https://github.com/coderpros)

# KickBox-ASP

A [KickBox.io](https://kickbox.io) API wrapper for Classic ASP.

## How to use

- Create a free account at [Kickbox.io](https://kickbox.io).
- Sign up for an API Key.
- Update the KICKBOX_API_KEY constant in ~/Includes/config.asp file to use your API Key.

### Verify a single email address

```vbs
Dim KickBox: Set KickBox = New KickBoxObject
Dim verificationResponse

With KickBox
    .ClientApiKey = KICKBOX_API_KEY
    Set verificationResponse = .VerifySingleEmail(string emailAddress))
End With
```

### Verify multiple email addresses

```vbs
Dim KickBox: Set KickBox = New KickBoxObject
Dim verificationResponse

With KickBox
    .ClientApiKey = KICKBOX_API_KEY
    Set verificationResponse = .VerifyBulkEmail(string emailAddressesSeparatedByNewline))
End With
```

### Check status of a bulk verification job

```vbs
Dim KickBox: Set KickBox = New KickBoxObject
Dim verificationResponse

With KickBox
    .ClientApiKey = KICKBOX_API_KEY
    Set verificationResponse = .CheckVerificationStatus(int JobId)
End With
```

*~/Default.asp is a functional example of how to use this library.*


[contributors-shield]: https://img.shields.io/github/contributors/coderpros/kickbox-asp.svg?style=flat-square
[contributors-url]: https://github.com/coderpros/kickbox-asp/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/coderpros/kickbox-asp?style=flat-square
[forks-url]: https://github.com/coderpros/kickbox-asp/network/members
[stars-shield]: https://img.shields.io/github/stars/coderpros/kickbox-asp.svg?style=flat-square
[stars-url]: https://github.com/coderpros/kickbox-asp/stargazers
[issues-shield]: https://img.shields.io/github/issues/coderpros/kickbox-asp?style=flat-square
[issues-url]: https://github.com/coderpros/kickbox-asp/issues
[license-shield]: https://img.shields.io/github/license/coderpros/kickbox-asp?style=flat-square
[license-url]: https://github.com/coderpros/kickbox-asp/master/blog/LICENSE
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=flat-square&logo=linkedin&colorB=555
[linkedin-url]: https://linkedin.com/company/coderpros
[twitter-shield]: https://img.shields.io/twitter/follow/coderpronet?style=social
[twitter-follow-url]: https://img.shields.io/twitter/follow/coderpronet?style=social
[github-shield]: https://img.shields.io/github/followers/coderpros?label=Follow&style=social
[github-follow-url]: https://img.shields.io/twitter/follow/coderpronet?style=social
