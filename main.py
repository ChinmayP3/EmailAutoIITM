import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from email.mime.base import MIMEBase
import argparse
import csv

parser = argparse.ArgumentParser()
parser.add_argument('-u', '--user',required=True)
parser.add_argument('-p', '--password',required=True)
parser.add_argument('-pdf', '--pdf')
parser.add_argument('-f', '--file',required=True)

args = parser.parse_args()

# https://myaccount.google.com/lesssecureapps

email_address = args.user
password = args.password
smtp_server = 'smtp.gmail.com'

port = 587


def getHtmlContent(name,company):
    html_content = """
            <div dir="ltr">
            <p><font face="georgia, serif">Hello {},</font></p>
            <p><b>Greetings from the Placement and Internship Team of IIT Madras!</b></p>
            <p>We are writing this to you with reference to the prospects of recruitment from IIT Madras for the Placement season 2023-24. We cordially invite <b>{}</b> to participate in this upcoming placement session to further strengthen our relationship.</p>
            <p>Upon receiving your confirmation, we will send an official invitation with further details regarding the process soon. You can also reach out at the below-mentioned contact for any query.</p>
            <p>Also, the following table gives a brief overview of some of our labs with their research areas:</p>
            <table cellspacing="0" cellpadding="0" dir="ltr" border="1" style="width:0px;table-layout:fixed;font-size:10pt;font-family:arial;border-collapse:collapse;border:none">
            <colgroup>
                <col width="223">
                <col width="242">
            </colgroup>
            <tbody>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(139,195,74);font-weight:bold;text-align:center;border:1px solid rgb(204,204,204)">Some of the Labs</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(139,195,74);font-weight:bold;text-align:center;border:1px solid rgb(204,204,204)">Research Areas</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Interactive Intelligence Laboratory</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Data Mining, Machine Learning</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Speech and Music Technology Lab</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Speech processing,<br>Computational Neuroscience</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Computer Vision Lab</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Image Processing</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(238,247,227);text-align:center;border:1px solid rgb(204,204,204)">VLSI Laboratory</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(238,247,227);text-align:center;border:1px solid rgb(204,204,204)">VLSI Design,<br>Cryptography<br>Network Security</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(238,247,227);text-align:center;border:1px solid rgb(204,204,204)">Application-Specific Hardware Design Lab</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(238,247,227);text-align:center;border:1px solid rgb(204,204,204)">High-Performance Computing<br>Parallelization</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">High-Performance Computing &amp; Networking Lab</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Computer Networks,<br>Distributed Systems</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Computer Architecture and Systems Lab</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Computer Architecture<br>High-Performance Computing</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(238,247,227);text-align:center;border:1px solid rgb(204,204,204)">Algorithms and Complexity Theory Lab</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(238,247,227);text-align:center;border:1px solid rgb(204,204,204)">Computational Complexity Theory<br>Design And Analysis of Algorithms</td>
                </tr>
            </tbody>
            </table>
            <p>Please let us know if you need any further information from our side. We are looking forward to your reply.</p>
            <p>Thanks &amp; Regards,</p>
            <p>
                <font color="#444444"><b>Chinmay Anil Pranjale</b></font><br>
                <font color="#3d85c6">
                    Placement and Internship Coordinator, CSE MTech<br>
                    Indian Institute Of Technology Madras<br>
                    <a href="mailto:mtech.pc5@smail.iitm.ac.in" target="_blank">mtech.pc5@smail.iitm.ac.in</a><br>
                    +91 9834419974
                </font>
            </p>
            </div>
        """.format(name,company)
    return html_content


def getHtmlContent2(name,company):
    html_content = """
            <div dir="ltr">
            <p><font face="georgia, serif">Hello {},</font></p>
            <p><b>Greetings from the Placement and Internship Team of IIT Madras!</b></p>
            <p>We are writing this to you with reference to the prospects of recruitment from IIT Madras for the Internship season 2023-24. We cordially invite <b>{}</b> to participate in this internship session to further strengthen our relationship.</p>
            <p>Upon receiving your confirmation, we will send an official invitation with further details regarding the process soon. You can also reach out at the below-mentioned contact for any query.</p>
            <p>Also, the following table gives a brief overview of some of our labs with their research areas:</p>
            <table cellspacing="0" cellpadding="0" dir="ltr" border="1" style="width:0px;table-layout:fixed;font-size:10pt;font-family:arial;border-collapse:collapse;border:none">
            <colgroup>
                <col width="223">
                <col width="242">
            </colgroup>
            <tbody>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(139,195,74);font-weight:bold;text-align:center;border:1px solid rgb(204,204,204)">Some of the Labs</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(139,195,74);font-weight:bold;text-align:center;border:1px solid rgb(204,204,204)">Research Areas</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Interactive Intelligence Laboratory</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Data Mining, Machine Learning</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Speech and Music Technology Lab</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Speech processing,<br>Computational Neuroscience</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Computer Vision Lab</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Image Processing</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(238,247,227);text-align:center;border:1px solid rgb(204,204,204)">VLSI Laboratory</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(238,247,227);text-align:center;border:1px solid rgb(204,204,204)">VLSI Design,<br>Cryptography<br>Network Security</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(238,247,227);text-align:center;border:1px solid rgb(204,204,204)">Application-Specific Hardware Design Lab</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(238,247,227);text-align:center;border:1px solid rgb(204,204,204)">High-Performance Computing<br>Parallelization</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">High-Performance Computing &amp; Networking Lab</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Computer Networks,<br>Distributed Systems</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Computer Architecture and Systems Lab</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;text-align:center;border:1px solid rgb(204,204,204)">Computer Architecture<br>High-Performance Computing</td>
                </tr>
                <tr style="height:21px">
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(238,247,227);text-align:center;border:1px solid rgb(204,204,204)">Algorithms and Complexity Theory Lab</td>
                <td style="font-family:georgia;padding:2px 3px;vertical-align:bottom;background-color:rgb(238,247,227);text-align:center;border:1px solid rgb(204,204,204)">Computational Complexity Theory<br>Design And Analysis of Algorithms</td>
                </tr>
            </tbody>
            </table>
            <p>Please let us know if you need any further information from our side. We are looking forward to your reply.</p>
            <p>Thanks &amp; Regards,</p>
            <p><font color="#444444"><b>Chinmay Anil Pranjale</b></font></p>
            <p><font color="#3d85c6">Placement and Internship Coordinator, CSE MTech</font></p>
            <p><font color="#3d85c6">Indian Institute Of Technology Madras</font></p>
            <p><font color="#3d85c6"><a href="mailto:mtech.pc5@smail.iitm.ac.in" target="_blank">mtech.pc5@smail.iitm.ac.in</a></font></p>
            <p><font color="#3d85c6">+91 9834419974</font></p>
            </div>
        """.format(name,company)
    return html_content

if(args.pdf):
    pdf_file_path = args.pdf
    with open(pdf_file_path, 'rb') as pdf_file:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(pdf_file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{pdf_file_path}"')

first = True
with open(args.file, mode ='r')as file:
    csvFile = csv.reader(file)
    for lines in csvFile:
        if(first):
            first = False
            continue
        company = lines[0].upper().strip()
        mail = lines[1]
        try:
            msg = MIMEMultipart()
            if(args.pdf):
                msg.attach(part)
            msg['From'] = "CSE Placement Team IIT Madras<"+email_address+">"
            msg['Subject'] = 'IIT Madras Campus Recruitment Invitation for Internship and Placement 2023-24'

            msg['To'] = mail
            msg.attach(MIMEText(getHtmlContent("Team",company), 'html'))
            server = smtplib.SMTP(smtp_server, port)
            server.starttls()
            server.login(email_address, password)
            server.sendmail(email_address, msg['To'], msg.as_string())
            print("Email sent successfully to "+msg['To'])
        except Exception as e:
            print("Error: Unable to send email.")
            print(e)
        finally:
            server.quit()
