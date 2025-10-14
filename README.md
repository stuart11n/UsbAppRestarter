Coded with Gemini to scratch an itch. There are no bells and whistles.

<img width="848" height="635" alt="image" src="https://github.com/user-attachments/assets/becd88b8-69e6-4e93-be19-5c2e7d1863b5" />

Both SPAD.next and FFBeast Commander have issues if a USB device disappears/reappears due to connection after app start, USB glitches, device restart etc. This solves that by restarting specified applications when any USB device is connected.

When a USB device is connected, the specified applications are restarted. If they are not running, they will be started.

It can be configured to start on login and it minimizes to the tray.

It's a bit rough and ready, but works for me. 

Will likely add an icon, a USB device name filter in future, and option to only restart running applications. Though it'd be nice if SPAD.next and Commander could render this obsolete.

