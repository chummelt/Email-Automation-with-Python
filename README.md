# Email-Automation-with-Python

Welcome to an easy and effective way in automating a bulk of mails via Outlook with Python! There are several different requests in sending mass emails for different purposes. For some it might be relevant to include an attachment file, others only need to include personalized links.
The background of **this** project was to sent out mass emails for a research project regarding COVID-19. For that, I had to find a way to:  

- sent a **bulk of emails** (> 2,500) we gained via webscraping
- use **individualized links** in the mail, so each participant could upload their individual data
- **include a certificate file** (pdf) in the attachment
- **include html links** of several sites about the research project
- **include different graphics** (png) like the institute logo and examples of how the data should look like. 
- avoid the **sending limit** to not be identified as spam

Requirement for this task: 

- install Microsoft Outlook
- install Python extensions for Microsoft Windows Provides access to much of the Win32 API: pywin32

```
pip install pywin32
```
