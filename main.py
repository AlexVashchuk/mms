from config import *
import pandas as pd
import smtplib
import ssl


# открываем соединение, собираем и отправляем сообщение(make_message), возвращаем результат для логов
def send_mail(name, receiver, var_content):
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
        server.login(sender, password)
        try:
            make_message(server, receiver, var_content)
            print('message sent')
            log = ('письмо отправлено', name, receiver, var_content)
        except smtplib.SMTPException:
            print('Error: unable to send email to', receiver)
            log = ('Error: письмо не было отправлено', name, receiver, var_content)
        finally:
            server.quit()
            return log


# собираем и отправляем сообщение. Возвращаем ответ сервера.
def make_message(server, to_address, var_content):
    bcc = []
    from_addr = sender
    message_subject = "Код для получения баллов на портале НМО"
    message_begin = 'Добрый день! \n\nВы присутствовали 26.04.2022г. на конференции «Передовые технологии ' \
                    'в эндокринологии, эндокринной хирургии и онкологии» ' \
                    '\nОтправляю вам код для получения баллов на портале НМО \n\n'
    message_end = '\n\nНе забудьте его зарегистрировать \n\nС уважением Евгения!'
    message = "From: %s\r\n" % from_addr + \
              "To: %s\r\n" % to_address + \
              "Subject: %s\r\n" % message_subject + \
              "\r\n" \
              + message_begin + var_content + message_end
    receivers = [to_address] + bcc
    return server.sendmail(sender, receivers, message)


# итерируемся по файлу excel и отправляем сообщение для каждой строки
df = pd.read_excel(r'source\НМО.xlsx')
count = 0
logs = []
for i, raw in df.iterrows():
    logs.append(send_mail(raw[0], raw[1], raw[2]))
    count += 1
print(count, 'messages processed \n')

# пишем лог.txt
f = open(r'source\log.txt', 'w')
for i in logs:
    a, b, c, d = i
    f.write(a + ' ' + b + ' ' + c + ' ' + d + '\n')
f.close()

print('for more info see log.txt')
print("well done!")
