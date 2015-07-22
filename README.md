# Centralised Outlook Signature
## Summary
In progress :-)

## Setup
### Infrastructure requirements
The Add-in is designed to fetch information about the currently logged on user from Microsoft® Active Directory. A working environment with ADDS and properly domain-joined machines is necessary. Also a network share on an arbitrary file server is needed to host the configuration file and signature templates. Read-only access is perfectly sufficient for Domain Users/Authorized Users.

### Clients/Workstations
The following additional software packages are required on all client machines:
* [Microsoft® .NET Framework 4.5.1](http://www.microsoft.com/de-at/download/details.aspx?id=40779) (or newer)
* Compatible version of Microsoft® Outlook (see **Compatibility**)

## Compatibility
The Add-in was tested successfully with these versions of Microsoft® Outlook:
* 2007 32-Bit
* 2010 32-Bit
* 2013 32-Bit
* 2016-Pre 32-Bit
 
## Remarks
* The code isn't properly localized yet, display language is currently **German** only.

## 3rd party credits
This Add-in wouldn't be possible without the following (awesome) projects:
* [NetOffice © SebastianDotNet](http://netoffice.codeplex.com/)
* [Libarius © Benjamin Höglinger](https://github.com/nefarius/Libarius)
* [ReactiveProtobuf © Benjamin Höglinger](https://github.com/nefarius/ReactiveProtobuf)
* [ReactiveSockets © Clarius Labs](https://github.com/clariuslabs/reactivesockets)
* [log4net © Apache Software Foundation](https://logging.apache.org/log4net/)
* [CS-Script - The C# Script Engine © Oleg Shilo](http://www.csscript.net/)
* [INI File Parser © Ricardo Amores Hernández](https://github.com/rickyah/ini-parser)
