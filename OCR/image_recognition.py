from operator import iconcat

from PIL import Image
import time
import subprocess
from pytesseract import pytesseract
from com.dtmilano.android.viewclient import ViewClient

icon_5g = '5G'
subprocess.call("adb shell input keyevent KEYCODE_HOME")
device, serialno = ViewClient.connectToDeviceOrExit()
vc = ViewClient(device=device, serialno=serialno)
# vc.dump()
# device.wake()
# device.press('KEYCODE_HOME')
device.drag((10, 10), (5000, 5000), 100)
time.sleep(3)
subprocess.call("adb devices", shell=True)
subprocess.call('adb exec-out screencap -p > D:\\5G\Screenshots\screen.png', shell=True)

path_to_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
image_path = r"D:\\5G\Screenshots\screen.png"

# Opening the image & storing it in an image object
img = Image.open(image_path)

# Providing the tesseract executable
# location to pytesseract library
pytesseract.tesseract_cmd = path_to_tesseract

# Passing the image object to image_to_string() function
# This function will extract the text from the image
text = pytesseract.image_to_string(img)

# Displaying the extracted text
status_bar_icon = str(text).splitlines()[0]
print(status_bar_icon)

if str(status_bar_icon) == icon_5g:
    print(icon_5g + ' icon is present')
else:
    print(icon_5g + ' icon is not present')
