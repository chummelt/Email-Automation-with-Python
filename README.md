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
At first, we need to make sure that we use the right wd. We need to create a folder which contains all used files like the csv with contact info and the images. We also need to import the win32com library. 
Now, import ```csv``` for the contact information. Since we also want to sent out images, we need to import ```pathlib```, too. 

```ruby
import os
os.chdir("C:/Users/.../Mail")
import csv
from time import sleep
import win32com.client as client
import pathlib
```

When using Outlook attachment, we have to use the absolute path. In this case, we are attaching a PDF called 'Ethicvotum'. Later on, we attach this PDF to our massage with ```message.Attachments.Add(#_absolute)```. For structural reasons, this is included in the main part of the code. 

```ruby
ethic_path = pathlib.Path('Ethicvotum.pdf')
ethic_absolut = str(ethic_path.absolut())
```
Now, we use the **csv** file 'contact' that contains the email addresses and personalized links. If you have the additional name in the **csv** file, this would work the same way. 
```ruby
# open distribution list
with open('contact.csv', 'r', newline='') as f:
    reader = csv.reader(f)
    distro = [row for row in reader] # (email, links)
```
Dispatch an instance of Outlook. 
```ruby
# create outlook instance
outlook = client.Dispatch('Outlook.Application')
```
To avoid landing in the spam filter or being limited with a maximum, the emails are chunked into blocks of 30 and then sent at a rate of 30 per minute by using a timer.
```ruby
# chunk distribution list into blocks of 30
chunks = [distro[x:x+30] for x in range(0, len(distro), 30)]
```
The message recipients are set with the ```To```, ```CC```, and ```BCC``` properties. In this example, we only use ```To```. Thus, we tell Outlook to sent message to 'address'. You could easily switch 'links' with 'names', if you do not have individual links but rather want to address the individual names. With ```Subject``` we can name the subject line which will be displayed in the mail. As mentioned earlier, **now** we include the command for the ```Attachment``` for the PDF file we inserted before. 
```ruby
# iterate through chunks and send mail
for chunk in chunks:
    # iterate through each recipient in chunk and send mail
    for link, address in chunk:
        message = outlook.CreateItem(0)
        message.To = address
        message.Subject = "Study about COVID-19"
        message.Attachments.Add(ethic_absolut
```
Now, we add the content of the mail in HTML format. You can use everything you want in HTML format, including images, changes of font and so on. 
```ruby
        html_body = """   
<pre><span class="pl-c">This is a very simple example, which includes an image from wikipedia.<br />
Do not forget to include your links, names or any other information you would want to include with '<span style="color: #339966;"><strong>{}</strong></span>'.<br />
Keep in mind that the placeholder for the link is now once used.
So, you would need to use it once again or adjust <br />
the <strong>'html_body.format(link,link)</strong>' to <strong>'html_body.format(link)'</strong>.<br />
<br />
Good luck and have fun!<br />
<br />
<img src="https://upload.wikimedia.org/wikipedia/commons/thumb/9/9c/Fender_Jazz-Bass_1966.jpg/320px-Fender_Jazz-Bass_1966.jpg" alt="Jazz Bass" width="320" height="940" /><br /></span></pre>    """

```
Now, each time we want to insert the individualized link (or individual name), include ```{}``` within the HTML string. 
It is important to check the **order** in which you want to want to refer to the individual information. 

```ruby
       message.HTMLBody = html_body
       message.HTMLBody = html_body.format(link,link)
       message.Send()
```
By telling the  argument ```html_body.format(link,link)```, the first set of strings from the **csv**, which includes the links, is inserted here. If we had a **csv** file which included 'names, addresses', we could have stated that with  ```for link, address in chunk:``` and use ```html_body.format(name,link)``` for example. 
We must not forget the timer for the chunking. Setting the timer to 60 seconds lets Outlook work stable enough and every mail gets out.

```ruby
    # wait 60 seconds before sending next chunk
    sleep(60)
 ```
