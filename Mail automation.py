###Email Automation with Python###

import os
os.chdir("C:/Users/.../Mail")
import csv
from time import sleep
import win32com.client as client
import pathlib

ethic_path = pathlib.Path('Ethicvotum.pdf')
ethic_absolute = str(ethic_path.absolute())


# open distribution list
with open('contact.csv', 'r', newline='') as f:
    reader = csv.reader(f)
    distro = [row for row in reader]

# chunk distribution list into blocks of 30
chunks = [distro[x:x+30] for x in range(0, len(distro), 30)]

# create outlook instance
outlook = client.Dispatch('Outlook.Application')

# iterate through chunks and send mail
for chunk in chunks:
    # iterate through each recipient in chunk and send mail
    for link, address in chunk:
        message = outlook.CreateItem(0)
        message.To = address
        message.Subject = "Research COVID-19"
        message.Attachments.Add(ethic_absolute)
        html_body = """
<pre><span class="pl-c">This is a very simple example, which includes an image from wikipedia.<br />Do not forget to include your links, names or any other information you would want to include with '<span style="color: #339966;"><strong>{}</strong></span>'.<br />Keep in mind that the placeholder for the link is now once used. So, you would need to use it once again or adjust <br />the <strong>'html_body.format(link,link)</strong>' to <strong>'html_body.format(link)'</strong>.<br /><br />Good luck and have fun!<br /><br /><img src="https://upload.wikimedia.org/wikipedia/commons/thumb/9/9c/Fender_Jazz-Bass_1966.jpg/320px-Fender_Jazz-Bass_1966.jpg" alt="Jazz Bass" width="320" height="940" /><br /></span></pre>    """
        message.HTMLBody = html_body
        message.HTMLBody = html_body.format(link,link)
        message.Send()

    # wait 60 seconds before sending next chunk
    sleep(60)
