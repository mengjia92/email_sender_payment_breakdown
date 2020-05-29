import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# display all columns
pd.set_option('display.max_columns', None)
pd.options.display.max_colwidth = 1000


class EmailSender:
    def __init__(self, file, usr, pwd, month, period):
        self.file = file
        self.usr = usr
        self.pwd = pwd
        self.month = month
        self.period = period

    def process_file(self):
        """
        read the raw data csv file, formatting the values inside
        :return: data in object type, grouped by tutor ID
        """
        data = pd.read_csv(self.file)
        data['Final Pay'] = data['Final Pay'].str.replace('$', '').astype(float)
        data['Object ID'] = data['Object ID'].fillna(0).astype(int)
        data = data.fillna('')
        # print(data.dtypes)
        return data.groupby('Tutor ID')

    def send_email(self, tutor_data, tutor_email, tutor_name, tutor_id):
        """
        Email sending module
        :param tutor_data:
        :param tutor_email:
        :param tutor_name:
        :param tutor_id:
        :return: None
        """
        print("Sending to " + str(tutor_id) + " " + tutor_name + "...")

        cols = ['Payment ID', 'Service', 'Payment Type', 'Incident', 'Object ID', 'Service Provider', 'Final Pay',
                'Currency', 'Approval Status', 'Admin Reason', 'Tutor Reason']
        msg = MIMEMultipart('related')

        msg['Subject'] = "[Easyke] " + self.month + " Payment Breakdown"
        msg['From'] = self.usr
        msg['To'] = tutor_email
        html_text = """\
                <html>
                  <head></head>
                  <body>
                    <p>Dear {0} (ID: {1}),</p>
                    <p>Thank you for teaching at XXX! You have earned a grand total of <b>${2}</b> in {3}.</p>
                    <p>Below is the payment breakdown for {4}:
                        {5}
                    </p>
                    <p>Regards,
                        <br>Jane Doe</br>
                    </p>
                  </body>
                </html>

            """.format(tutor_name, tutor_id, tutor_data['Final Pay'].agg('sum'), self.month, self.period,
                       tutor_data[cols].to_html())

        part_html = MIMEText(html_text, 'html')
        msg.attach(part_html)

        try:
            # ---------- SMTP TLS ----------
            # ser = smtplib.SMTP('smtp.gmail.com', 587)
            # ser.starttls()
            # ser.ehlo()
            # ---------- SMTP SSL ----------
            ser = smtplib.SMTP_SSL('smtp.gmail.com', 465)

            ser.login(self.usr, self.pwd)
            ser.sendmail(self.usr, tutor_email, msg.as_string())
            # ser.quit()
            print("Success!")
        except smtplib.SMTPServerDisconnected as info:
            print("Unable to send, " + str(info))


def main():
    es = EmailSender(data_file, username, password, month, period)
    all_tutors = es.process_file()

    for ID, tutor in all_tutors:
        # if ID > 3650:
        es.send_email(tutor, tutor['Tutor Email'].unique()[0], tutor['Tutor Name'].unique()[0], ID)

    # ---------- Send individually when SMTP error occurs ----------
    # t_id = 3650
    # t = all_tutors.get_group(t_id)
    # es.send_mail(t, t['Tutor Email'].unique()[0], t['Tutor Name'].unique()[0], t_id)


if __name__ == '__main__':
    data_file = 'raw_data.csv'
    username = 'XXX@gmail.com'
    password = 'xxxxx'
    month = 'April'
    period = 'Apr. 1 - Apr. 30'

    main()