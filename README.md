Some applications have issues if a USB device disappears/reappears due to connection after app start due to USB glitches, device restart, plugging in etc. This solves that by restarting specified applications when any USB device is connected. 

Here it is configured to restart SPAD.next and FFbeast Commander when my FFBeast yoke connects:

<img width="793" height="801" alt="image" src="https://github.com/user-attachments/assets/75cc208c-3640-4885-a315-35c9b8fc4e92" />

When a USB device is connected, the specified applications are restarted. If they are not running, they will be started. 
If you want a general restart on _any_ USB connection event add "USB" as a filter. It can be configured to start on login and it minimizes to the tray. 

It's a bit rough and ready, but works for me. It'd be nice if SPAD.next and Commander could render this obsolete. USBTreeView is helpful for finding the USB ID of devices. 

First 80% coded with Gemini 🐐 the rest is is 100% pure, old-fashioned, homegrown human. 
