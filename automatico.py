import schedule
import time
import globals
from envio import enviarEmail

schedule.every().day.at("14:42").do(enviarEmail)

while globals.loop:
    schedule.run_pending()
    time.sleep(60)

