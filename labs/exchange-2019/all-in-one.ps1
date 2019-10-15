#
<#
.SYNOPSIS
    Script to install all the needed pre-reqs for Exchange 2019 and a Domain Controller
.DESCRIPTION
    This script is intended to be used during the build process for an all-in-one lab in Azure DevTest
.NOTES
    Authored By: David Milward
    Email: support@exclaimer.com
    Date: 15th October 2019
.HISTORY
    1.0 - Installs Exchange pre-reqs and DC features
#>

Install-WindowsFeature -Name Install-WindowsFeature Web-WebServer,Web-Common-Http,Web-Default-Doc,Web-Dir-Browsing,Web-Http-Errors,Web-Static-Content,Web-Http-Redirect,Web-Health,Web-Http-Logging,Web-Log-Libraries,Web-Request-Monitor,Web-Http-Tracing,Web-Performance,Web-Stat-Compression,Web-Dyn-Compression,Web-Security,Web-Filtering,Web-Basic-Auth,Web-Client-Auth,Web-Digest-Auth,Web-Windows-Auth,Web-App-Dev,Web-Net-Ext45,Web-Asp-Net45,Web-ISAPI-Ext,Web-ISAPI-Filter,Web-Mgmt-Tools,Web-Mgmt-Compat,Web-Metabase,Web-WMI,Web-Mgmt-Service,NET-Framework-45-ASPNET,NET-WCF-HTTP-Activation45,NET-WCF-MSMQ-Activation45,NET-WCF-Pipe-Activation45,NET-WCF-TCP-Activation45,Server-Media-Foundation,MSMQ-Services,MSMQ-Server,RSAT-Feature-Tools,RSAT-Clustering,RSAT-Clustering-PowerShell,RSAT-Clustering-CmdInterface,RPC-over-HTTP-Proxy,WAS-Process-Model,WAS-Config-APIs
Install-WindowsFeature -Name AD-Domain-Services -IncludeManagementTools