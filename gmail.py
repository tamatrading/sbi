# メール送信
import smtplib
from email.mime.text import MIMEText
from email.utils import formatdate

#-----------------------------
# gmail送信
#-----------------------------
def sendGmail(fromAddress, toAddress, sendAddress, password, subject, bodyText):

    # SMTPサーバに接続
    smtpobj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpobj.starttls()
    smtpobj.login(sendAddress, password)

    # メール作成
    msg = MIMEText(bodyText)
    msg['Subject'] = subject
    msg['From'] = fromAddress
    msg['To'] = toAddress
    msg['Date'] = formatdate()

    # 作成したメールを送信
    smtpobj.send_message(msg)
    smtpobj.close()

