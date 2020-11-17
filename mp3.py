# Python supports COM, if you have the Win32 extensions
# check your Python folder eg.  D:\Python23\Lib\site-packages\win32com
# also http://starship.python.net/crew/mhammond/win32/Downloads.html
# this program will play MP3, WMA, MID, WAV files via the WindowsMediaPlayer
from win32com.client import Dispatch
mp = Dispatch("WMPlayer.OCX")
# use an mp3 file you have ...
#tune = mp.newMedia("C:/Program Files/Common Files/HP/Memories Disc/2.0/audio/Swing.mp3")
# or copy one to the working folder ...
tune = mp.newMedia("bgm.mp3")
# you can also play wma files, this cool sound came with XP ...
# tune = mp.newMedia("C:/WINDOWS/system32/oobe/images/title.wma")
mp.currentPlaylist.appendItem(tune)
mp.controls.play()
# to stop playing use
raw_input("Press Enter to stop playing")
mp.controls.stop()
# http://www.sharejs.com/codes/python/5733
