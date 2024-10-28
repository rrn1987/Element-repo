import sys
import os
import time
import re

Dump='vc.dump()'
Wifi_Preffered ='Wi-Fi preferred'
Cellular_Preffered ='Cellular preferred'

def call_component(package, activity):
    component = package + "/" + activity
    if True:
     device.startActivity(component=component)

def select_WiFi_Preferrred(preffered_type):
    device.startActivity('com.android.settings/.Settings')
    vc.dump()
    vc.findViewByIdOrRaise("com.android.settings:id/main_content_scrollable_container").uiScrollable.flingToBeginning(1)
    vc.dump()
    vc.findViewWithText('Network & internet').touch()
    vc.dump()
    if vc.findViewWithText('Mobile network'):
        vc.findViewWithText('Mobile network').touch()
    elif vc.findViewWithText('SIMs'):
        vc.findViewWithText('SIMs').touch()
    vc.dump()
    vc.findViewWithText('Advanced').touch()
    vc.dump()
    wifi_calling_text = 'Wi-Fi calling'
    print(wifi_calling_text)
    vc.findViewWithTextOrRaise(wifi_calling_text).touch()
    time.sleep(2)
    vc.dump()
    vc.findViewWithText('Wi-Fi calling').touch()
    vc.dump()
    text_off = vc.findViewWithText('Off')
    if text_off:
        text_off.touch()
        time.sleep(2)
    vc.dump()
    vc.findViewWithText('Calling preference').touch()
    vc.dump()
    vc.findViewWithText(Wifi_Preffered).touch()
    time.sleep(2)
    device.press('KEYCODE_BACK')
    device.press('KEYCODE_HOME')

if __name__ == '__main__':
    from com.dtmilano.android.viewclient import ViewClient
    device, serialno = ViewClient.connectToDeviceOrExit()
    vc = ViewClient(device=device, serialno=serialno)
    vc.dump()
    # device.openQuickSettings()
    # device.wake()
    # device.press('KEYCODE_HOME')
    # device.drag((10, 10), (5000, 5000), 100)
    # vc.dump()
    # print(vc.findViewByIdOrRaise("com.android.systemui:id/mobile_combo").getText())
    # ViewClient(*ViewClient.connectToDeviceOrExit()).traverse()
    # select_WiFi_Preferrred(Wifi_Preffered)
    # device.startActivity('com.android.settings/.Settings')
    # # vc.device.shell("svc wifi enable")
    # vc.dump()
    # vc.findViewWithText('Network & internet').touch()
    # vc.dump()
    # vc.findViewWithText('Wi‑Fi').touch()

    ViewClient(*ViewClient.connectToDeviceOrExit()).traverse()
    device.shell('am start com.google.android.apps.messaging')
    vc.dump()
    vc.findViewById('com.google.android.apps.messaging:id/start_new_conversation_button').touch()

