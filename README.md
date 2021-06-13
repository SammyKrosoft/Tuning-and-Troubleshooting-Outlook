# Tuning and Troubleshooting Outlook
> By **Wade Smith** – Awesome Microsoft Engineer

#### 1.	Outlook data files can not be scanned by AV software.  Any OST or PST file should not be scanned.

[Planning considerations for deploying Outlook for Windows](https://docs.microsoft.com/en-us/outlook/troubleshoot/deployment/plan-outlook-2016-deployment#outlook-security-considerations)

#### 2.	Verify mailbox sizes and ensure no one is approaching the 10GB mark.    Once an OST file hits 10GB, Outlook will start to experience minor performance delays and hangs.  If the OST approaches 20GB, performance issues will become more significant and Outlook may crash.

[You may experience application pauses if you have a large Outlook data file](https://support.microsoft.com/en-us/help/2759052/you-may-experience-application-pauses-if-you-have-a-large-outlook-data)

#### 3.	Microsoft provides recommendations on the number of mail items in folders, the number of folders, the number of items in calendars, and the number of items in shared mailboxes that a user can have before performance will degrade.

[Outlook performance issues when there are too many items or folders in a Cached mode .ost or .pst file folder](https://support.microsoft.com/en-us/help/2768656/outlook-performance-issues-when-there-are-too-many-items-or-folders-in)

Exchange administrators can run reports on the number of mail items in user mailboxes via PowerShell.

[Get-MailboxFolderStatistics](https://docs.microsoft.com/en-us/powershell/module/exchange/get-mailboxfolderstatistics?view=exchange-ps)

#### 4.	Outlook cache mode settings allow you to specify how much data from the mailbox will be downloaded.  This setting is available to users.  The recommendation is to permit 3-12 months of data to be downloaded depending on your organizational requirements.

#### 5.	Third party add-ins need to be reviewed in regards to Outlook performance and stability.  Add-in’s that are not required should be disabled, and the latest add-in from the manufacturer should be applied.

#### 6.	Network issues and network latency will cause poor performance in Outlook.  Microsoft has published recommendations on the latency in which Outlook can handle before the user experience will start to suffer.  The latency numbers can be viewed in the ‘Connection Status’ in Outlook.

To get Network Round Trip Time (RTT) you need to do (Avg Resp - Avg Proc).  That will give you the Network RTT between client and server.

If it's below 325ms => great for OL cache mode
If it's below 110ms => great for OL Online mode

*Example:*

```              
Max Avg Proc Time (Exchange RPC Latency) = 25ms
Max Network RTT Time (Network Ping Time) = 300ms
Max Avg Resp Time (Exchange RPC Latency + Network Latency) = 325ms
```
> **Please note**: These numbers are just one part of the troubleshooting process that is used when looking at Outlook performance issues.  There are many other factors that need to be taken in consideration such as the Outlook version, the mailbox size, number of items in the mailbox, AV, 3rd party add-ins, etc..)

#### 7.	For mailboxes that have ‘Shared Mailboxes’ added into the profile, the shared mailboxes should also be in cached mode.  The less contact we need to have between the client and server, the less issues we will see due to network latency over the VPN.

#### 8.	There is a 64 bit version of Office, but the performance benefits for a 64 bit version of Outlook (compared to 32 bit) are minimal.  Excel, PowerPoint, Project and Access will see the most improvement in performance.  As well, you need to ensure any 3rd party add-ins support the 64 bit version of Office.

[Choose between the 64-bit or 32-bit version of Office](https://support.microsoft.com/en-us/office/choose-between-the-64-bit-or-32-bit-version-of-office-2dee7807-8f95-4d0c-b5fe-6c6f49b8d261)

#### 9.	The following tools\guides can be used when troubleshooting Outlook performance issues:

- [Microsoft Support and Recovery Assistant](https://www.microsoft.com/en-us/download/100607)

- [How to scan Outlook by using the Microsoft Support and Recovery Assistant](https://support.microsoft.com/en-ca/help/4098558/scan-outlook-by-using-microsoft-support-and-recovery-assistant)

- [Using Process Explorer to List dll’s Running Under the Outlook.exe Process](https://support.microsoft.com/en-us/help/970920/using-process-explorer-to-list-dlls-running-under-the-outlook-exe-proc)

- [How to enable global and advanced logging for Microsoft Outlook](https://support.microsoft.com/en-ca/help/2862843/how-to-enable-global-and-advanced-logging-for-microsoft-outlook)

- [How to troubleshoot performance issues in Outlook](https://support.microsoft.com/en-ca/help/2695805/how-to-troubleshoot-performance-issues-in-outlook)

- [Plan and configure Cached Exchange Mode in Outlook 2016 for Windows](https://docs.microsoft.com/en-us/outlook/troubleshoot/deployment/cached-exchange-mode)


#### 10. For Microsoft escalation engineers to review data traces, all 3rd party add-ins need to be unloaded\uninstalled and Outlook needs to be running in a supported configuration per all of the recommendations listed above.

#### 11. Microsoft has published a ‘Best Practices’ guide that was written by the ‘Microsoft Outlook Product Team’.  The guide is designed to highlight ‘Best Practices’ to provide the best experience in Outlook.

I find the guide may be a bit ‘heavy’ for an average user, so this is something that the administration teams can review and possibly provide some ‘tips and tricks’ for users.

[Best Practices for Outlook](https://support.microsoft.com/en-us/office/best-practices-for-outlook-f90e5f69-8832-4d89-95b3-bfdf76c82ef8)

#### 12. Microsoft also has a guide for ‘Best Practices’ when using the Outlook calendar.  (Would be good for clients who are delegates for other users.)

[Best practices for organizations when using the Outlook Calendar](https://support.microsoft.com/en-us/office/best-practices-for-organizations-when-using-the-outlook-calendar-d93f72d3-2361-4e0d-8d6a-5c4939c17f39?ui=en-us&rs=en-us&ad=us)

##  Conclusion

With all of the information listed above, this will be a good start to ‘maximize’ performance for Outlook.  Most of these settings apply to both on-premises Exchange and Office 365.  

Overall, all of the recommendations above will need to be reviewed\implemented to ensure your clients have a good ‘user experience’ in Outlook.  (Especially when connected via VPN)
