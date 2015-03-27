# monitor player events from MediaMonkey.  print(event flag and player state info for each
# note: once started, script does not exit until MM is shut down.
import mpd
import sys
import pythoncom
import win32com.client
import time
import json

_mpdserver = "192.168.1.11"
_mpdport = 6600
_pathfixfrom1= "D:\\Music"
_pathfixto1= "USB/music"
_pathfixfrom2= "\\"
_pathfixto2= "/"
_mpdplaylist="MusicServer Playlist"
_maxpath=256
boolReps = ['F', 'T']   # hacky!

_quiting = False
_mpdclient = None
_sdbclient = None

class MMEventHandlers():
    def __init__(self):
        self._play_events = 0
    def showMMStatus(self):
        # note: MMEventHandlers instance includes all of SDBApplication members as well
        print('MM: Play:', boolReps[_sdbclient.Player.isPlaying], '; Pause:', boolReps[_sdbclient.Player.isPaused], '; Song Index:', _sdbclient.Player.CurrentSongIndex, '|', _sdbclient.Player.PlaylistCount)
        if _sdbclient.Player.isPlaying:
            print('Current Song:', _sdbclient.Player.CurrentSong.Title[:_maxpath])
    def OnShutdown(self):   #OK
        global _quiting	
        print(">> MMEventHandlers.OnShutdown") 
        _quiting = True
    def OnPlay(self):       #OK
        self._play_events += 1
        print(">> MMEventHandlers.OnPlay")
        if _sdbclient.Player.CurrentSongIndex>-1:
            _mpdclient.play(_sdbclient.Player.CurrentSongIndex)
        self.showMMStatus()
    def OnPause(self):      #OK
        print(">> MMEventHandlers.OnPause")
        if _sdbclient.Player.isPaused: 
            _mpdclient.pause(1)
        else:
            _mpdclient.pause(0)
        self.showMMStatus()
    def OnStop(self):
        print(">> MMEventHandlers.OnStop")
        if not _quiting:
            _mpdclient.stop()
        self.showMMStatus()
    def OnTrackEnd(self):
        print(">> MMEventHandlers.OnTrackEnd")
        self.showMMStatus()
    def OnPlaybackEnd(self):
        if not _quiting:
            print(">> MMEventHandlers.OnPlaybackEnd")
            self.showMMStatus()
    def OnCompletePlaybackEnd(self):
        print(">> MMEventHandlers.OnCompletePlaybackEnd")
        self.showMMStatus()
    def OnSeek(self):       #OK
        print(">> MMEventHandlers.OnSeek")
        print ('Seek [', _sdbclient.Player.PlaybackTime, '|', _sdbclient.Player.CurrentSongLength, ']')
        _mpdclient.seek(_sdbclient.Player.CurrentSongIndex, int(_sdbclient.Player.PlaybackTime*.001))
        self.showMMStatus()
    def OnNowPlayingModified(self):     #OK
        print(">> MMEventHandlers.OnNowPlayingModified")
        SyncMMNowPlayToMPD()
        self.showMMStatus()
    def OnTrackSkipped(self, track):  #OK (only when playing)
        print(">> MMEventHandlers.OnTrackSkipped")
        self.showMMStatus()
        # the type of any argument to an event is PyIDispatch
        # here, use PyIDispatch.Invoke() to query the 'Title' attribute for printing
        print('[', track.Invoke(3,0,2,True), ']')
    #def OnIdle(self):     #OK
        #print(">> OnIdle")
    def OnPlaylistAdded (self, playlist):
        print(">> MMEventHandlers.OnPlaylistAdded")
        print('[', playlist.Invoke(201,0,2,True), ']')
    def OnPlaylistDeleted (self, playlist):
        print(">> MMEventHandlers.OnPlaylistDeleted")
        print('[', playlist.Invoke(201,0,2,True), ']')
    def OnPlaylistChanged (self, playlist):
        print(">> MMEventHandlers.OnPlaylistChanged")
        print('[', playlist.Invoke(201,0,2,True), ']')

def MPDClearPlaylists():
    #Remove current playlist
    mpdplaylists=_mpdclient.listplaylists();
    for itm in mpdplaylists:
        _mpdclient.rm(itm['playlist'])
        
def SyncMMPlaylistToMPD():

    #Remove current playlist
    mpdplaylists=_mpdclient.listplaylists();
    for itm in mpdplaylists:
        _mpdclient.rm(itm['playlist'])
    
    #Find MPC Playlist in MM
    rootplaylist=_sdbclient.PlaylistByID(-1)
    playlists=rootplaylist.ChildPlaylists
    for i in range(0, playlists.Count):
        itm=playlists.Item(i)
        if itm.Title == _mpdplaylist:
            playlists=itm.ChildPlaylists

            #Iterate through MPC Child Playlist
            for j in range(0, playlists.Count):
                itm=playlists.Item(j)
                playlisttitle=itm.Title
                print (" ", playlisttitle)
                _mpdclient.clear();
                tracks=itm.Tracks

                #Iterate through Playlist Tracks
                for k in range(0, tracks.Count):
                    itm=tracks.Item(k)
                    mpdtrack=itm.Path
                    fixedmpdtrack=FixString (mpdtrack)
                    print ("   ", fixedmpdtrack)
                    try: 
                        _mpdclient.add(fixedmpdtrack);
                    except: 
                        pass 
                _mpdclient.save(playlisttitle)
            break;
            

#todo: stop if current song is removed  
def SyncMMNowPlayToMPD():
    print("?MM Count:" , _sdbclient.Player.CurrentSongList.Count)
    mpdcount=int(_mpdclient.status()['playlistlength'])
    print ("?MPD Count ", mpdcount)
    if _sdbclient.Player.CurrentSongList.Count == 0:
        # playlist cleared
        print ("Playlist Changed: Cleared")
        _mpdclient.clear()	
        _mpdclient.stop()
    else:
        if _sdbclient.Player.CurrentSongList.Count == mpdcount or _sdbclient.Player.CurrentSongList.Count > mpdcount:
             #songs added  or moved in playlist
            if mpdcount == 0:
                print ("Playlist Changed: MPD Playlist Empty")
                for i in range(0, _sdbclient.Player.CurrentSongList.Count):
                    mmsong=_sdbclient.Player.CurrentSongList.Item(i).Path[:_maxpath]
                    fixedmmsong=FixString (mmsong)
                    print("   + ", fixedmmsong)
                    _mpdclient.add(fixedmmsong)
            else:
                    print ("Playlist changed: Song added moved")
                    for mmindex in range(0, _sdbclient.Player.CurrentSongList.Count):
                        mmsong=_sdbclient.Player.CurrentSongList.Item(mmindex).Path[:128]
                        fixedmmsong=FixString (mmsong)
                        mpdcount=int(_mpdclient.status()['playlistlength'])
                        found=0
                        for mpdindex in range(mmindex, mpdcount):
                            mpdsong=_mpdclient.playlist()[mpdindex]
                            fixedmpdsong=mpdsong.replace("file: ", "")
                            if fixedmmsong == fixedmpdsong:
                                found=1
                                if mmindex == mpdindex:
                                    print ("Ok: ", fixedmmsong)
                                    break
                                else:
                                    print ("   m", fixedmmsong)
                                    _mpdclient.move(mpdindex, mmindex)
                        if not found:
                            _mpdclient.add(fixedmmsong)
                            print ("   +", fixedmmsong)
                            if mpdcount != mmindex:
                                _mpdclient.move(mpdcount, mmindex)
                                print ("   m", fixedmmsong)
        else:
            #songs removed from playlist
            print ("Playlist Changed: Songs Removed")
            i=0
            while i < mpdcount:
                mpdsong=_mpdclient.playlist()[i]
                fixedmpdsong=mpdsong.replace("file: ", "")
                if i <_sdbclient.Player.CurrentSongList.Count:
                    mmsong=_sdbclient.Player.CurrentSongList.Item(i).Path[:_maxpath]
                    fixedmmsong=FixString (mmsong)
                    if fixedmmsong == fixedmpdsong:
                        print ("Ok: ", fixedmmsong)
                        i=i+1
                    else:
                        _mpdclient.delete(i)
                        mpdcount=mpdcount-1
                        print("   -", fixedmpdsong)
                else:
                    _mpdclient.delete(i)
                    mpdcount=mpdcount-1
                    print("   -", fixedmpdsong)
        
    # Check sync
    syncerror=False
    if _sdbclient.Player.CurrentSongList.Count != int(_mpdclient.status()['playlistlength']):
        syncerror=True    
    else:
        for i in range(0, _sdbclient.Player.CurrentSongList.Count):
            mpdsong=_mpdclient.playlist()[i]
            fixedmpdsong=mpdsong.replace("file: ", "")
            mmsong=_sdbclient.Player.CurrentSongList.Item(i).Path[:_maxpath]
            fixedmmsong=FixString (mmsong)
            if fixedmmsong != fixedmpdsong:
                syncerror=True
    if syncerror:
        print("**** Sync Error **** ")

def FixString(string):
    fixedstring=string.replace(_pathfixfrom1, _pathfixto1)
    fixedstring=fixedstring.replace(_pathfixfrom2, _pathfixto2)
    return fixedstring

def StopMMMonitor():
    _quiting=True

def StartMMMonitor():

    # running the script will start MM if it's not already running
    global _sdbclient
    _sdbclient = win32com.client.DispatchWithEvents('SongsDB.SDBApplication', MMEventHandlers)
    print ("** monitor started ***")

    _mpdclient.stop()
    _mpdclient.clear()
    SyncMMNowPlayToMPD()

    while not _quiting:
        # required by this script because no other message loop running
        # if the app has its message loop (i.e., has a Windows UI), then
        # the events will arrive with no additional handling
        pythoncom.PumpWaitingMessages()
        time.sleep(0.2)
 
    # note that SDB instance includes members of of the MMEventHandlers class
    print ("** monitor stopped; received " + str(_sdbclient._play_events) + " play events ***")
 
def Main():
    global _sdbclient
    global _mpdclient
    _mpdclient = mpd.MPDClient(use_unicode=True)
    _mpdclient.connect(_mpdserver, _mpdport)

    total = len(sys.argv)
    cmdargs = str(sys.argv)
    if total == 1:
        #startMMMonitor()
        _sdbclient = win32com.client.Dispatch("SongsDB.SDBApplication")
        SyncMMPlaylistToMPD()
    else:
        for i in range(0,total):
            if str(sys.argv[i]) == '-startmonitor':
                StartMMMonitor()
            elif str(sys.argv[i]) == '-stopmonitor':
                StopMMMonitor()
            elif str(sys.argv[i]) == '-clear':
                _mpdclient.clear();
            elif str(sys.argv[i]) == '-play':
                _mpdclient.play();
            elif str(sys.argv[i]) == '-stop':
                _mpdclient.stop();
            elif str(sys.argv[i]) == '-pause':
                _mpdclient.pause();
            elif str(sys.argv[i]) == '-next':
                _mpdclient.next();
            elif str(sys.argv[i]) == '-previous':
                _mpdclient.previous();
            elif str(sys.argv[i]) == '-syncplaylists':
                _sdbclient = win32com.client.Dispatch("SongsDB.SDBApplication")
                SyncMMPlaylistToMPD()
            elif str(sys.argv[i]) == '-syncnowplaying':
                _sdbclient = win32com.client.Dispatch("SongsDB.SDBApplication")
                SyncMMNowPlayToMPD()

if __name__ == '__main__':
        Main()
