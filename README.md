# EXIF Remover

A very simple Windows .NET Framework application that is visible in tray. With it running you will get an Explorer option showing `Remove EXIF data`.

As you probably can guess, then this will remove ALL EXIF data from the image file, which is handy if you want to publish the file to the internet.

You can even drag/drop folders into the UI, and it will then scan all sub-folders for image files, and remove EXIF from it. **Be careful about this, as it will not ask for confirmation!**

# Screenshots

When dragging 4 files to the UI:
<img width="900" height="478" alt="image" src="https://github.com/user-attachments/assets/49e4f486-08c2-4a48-8a9e-142c0018103d" />

When selecting image files and using context menu:
<img width="900" height="315" alt="image" src="https://github.com/user-attachments/assets/af73de1e-4297-421a-b13e-25bc2b0a43ad" />

# Requirements

- Windows 10 or newer
- .NET Framework 4.8.1 (part of native OS)
