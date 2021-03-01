# MediaMonkey Addons

I manage my music and audiobook collection with MediaMonkey.
Here are some addons I made for myself, to manage tagging more easily.

* TagToolbox0.3.vbs: Makes adding and removing tags in batch much easier.
* showAllMetadata.vbs: Shows all the metadata of a file. It uses ExifTool by Phil Harvey to get the metadata from the files. Unfortunately the script is slow, because everytime a new file is highlighted it starts a separate exiftool.exe process, which takes a while.