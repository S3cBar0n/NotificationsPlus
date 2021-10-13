from threading import Thread, activeCount
from win10toast import ToastNotifier
from time import sleep
import win32com.client
import pythoncom
import datetime
import pytz

meeting_types = [0, 1, 2, 3]
meeting_response = {"0": "Accepted",
                    "1": "Organizer",
                    "2": "Tentative",
                    "3": "Accepted"}


def calendar_checker():
    # Initialized our COM object for Win32Com in a threaded application
    pythoncom.CoInitialize()
    # Toaster initialization for windows pop up notifications.
    toaster = ToastNotifier()
    # Date calculating variables
    today = datetime.datetime.today().strftime("%Y-%m-%d")
    now = datetime.datetime.now().astimezone(pytz.timezone("America/Denver")).strftime('%m/%d/%Y %H:%M')
    # Delta is to calculate meetings that will be starting within a 5 minute delta.
    delta = datetime.datetime.now().astimezone(pytz.timezone("America/Denver")) + datetime.timedelta(minutes=6)
    # Objects for Outlook, used to get our appointments
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendar = outlook.GetDefaultFolder("9")
    appointments = calendar.Items

    for appointment in appointments:
        # This grabs appointments that are starting today
        if today in str(appointment.Start):
            appt_time = appointment.Start.strftime('%m/%d/%Y %H:%M')
            # Calculating if the upcoming appointment has passed the current point in time.
            if (
                    now <= appt_time < delta.strftime('%m/%d/%Y %H:%M')
                    and appointment.ResponseStatus in meeting_types
            ):
                print(f"Calendar - {appointment.Subject}")
                toaster.show_toast(
                    f"MEETING STARTING SHORTLY - {meeting_response[str(appointment.ResponseStatus)]}",
                    f"{appointment.Subject}", duration=15, threaded=True)


# Loop for our thread that is getting started in main.
def calendar_checker_loop():
    print("Appointment Monitor Started")
    while True:
        sleep(150)
        try:
            calendar_checker()
        except Exception as e:
            print(f"Failed to perform Calendar Check: {e}")


def email_checker():
    # Initialized our COM object for Win32Com in a threaded application
    pythoncom.CoInitialize()
    # Toaster initialization for windows pop up notifications.
    toaster = ToastNotifier()
    # Objects for Outlook, used to get our emails
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder("6")
    messages = inbox.Items
    unread_msg = []

    for msg in messages:
        if msg.UnRead is True:
            unread_msg.append(f"{msg.SenderName} about {msg.Subject}")
            print(f"Email - {msg.SenderName} - {msg.Subject}")

    for subfolder in inbox.Folders:
        for folder in subfolder.Folders:
            messages = folder.Items
            for msg in messages:
                if msg.UnRead is True:
                    unread_msg.append(f"{msg.SenderName} about {msg.Subject}")
                    print(f"Email - {msg.SenderName} - {msg.Subject}")
        messages = subfolder.Items
        for msg in messages:
            if msg.UnRead is True:
                unread_msg.append(f"{msg.SenderName} about {msg.Subject}")
                print(f"Email - {msg.SenderName} - {msg.Subject}")

    if len(unread_msg) == 1:
        toaster.show_toast(f"Unread Message!!!", f"{unread_msg[0]}", duration=15, threaded=True)
    elif len(unread_msg) > 1:
        toaster.show_toast(f"{len(unread_msg)} Unread Messages!!!", '\n'.join(unread_msg), duration=15, threaded=True)


# Loop for our thread that is getting started in main.
def email_checker_loop():
    print("Email Monitor Started")
    while True:
        sleep(120)
        try:
            email_checker()
        except Exception as e:
            print(f"Failed to perform Calendar Check: {e}")


def main():
    print("Notifications Plus Successfully Initialized")
    # Starts a loop to run the calendar_checker function
    Thread(target=calendar_checker_loop).start()
    # Starts a loop to run the email_checker function
    Thread(target=email_checker_loop).start()

    # This is a safety measure to monitor our threads to ensure we do not overwhelm the hardware. Exits in the case
    # of more than 4 threads. It should never use more than 3 (1 thread for main, 2 threads for the loops)
    # However, an extra thread has been allocated for emergencies.
    print("Starting thread monitoring")
    while True:
        sleep(600)
        total_threads = activeCount()
        print(f"Checking active threads... {total_threads}")
        if total_threads > 4:
            exit()


if __name__ == '__main__':
    main()
