# Email-Automation-with-Python

Welcome to an easy and effective way in automating a bulk of mails via Outlook with Python! There are several different requests in sending mass emails for different purposes. For some it might be relevant to include an attachment file, others only need to include personalized links.
The background of **this** project was to sent out mass emails for a research project regarding COVID-19. For that, I had to find a way to:  

- sent a **bulk of emails** (> 2,500) we gained via webscraping
- use **individualized links** in the mail, so each participant could upload their individual data
- **include a certificate file** (pdf) in the attachment
- **include html links** of several sites about the research project
- **include different graphics** (png) like the institute logo and examples of how the data should look like. 
- avoid the **sending limit** to not be identified as spam

Therefore **this** project will include all of the above in the following. So, let's start!

Requirement for this task: 

- install Microsoft Outlook
- install Python extensions for Microsoft Windows Provides access to much of the Win32 API: pywin32

```ruby
pip install pywin32
```
At first, we need to make sure that we use the right wd. You need to create a folder which contains all used files like the csv with contact info and the images. We also need to import the win32com library and dispatch an instance of Outlook. We also need to import csv for the contact information. Since we also want to sent out images, we need ti import pathlib, too. 

```ruby
import os
os.chdir("C:/Users/.../Mail")
import csv
from time import sleep
import win32com.client as client
import pathlib
```

When using Outlook attachment, we have to use the absolute path. In this case, we are attaching a PDF calles 'Ethikvotum'. Later on, we attach this PDF to our massage with ```message.Attachments.Add(ethik_absolute)```. For structural reasons, this is included in the main part of the code. 

```ruby
ethik_path = pathlib.Path('Ethikvotum.pdf')
ethik_absolute = str(ethik_path.absolute())
```
