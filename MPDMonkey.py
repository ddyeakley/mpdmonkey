from mpd import MPDClient, MPDError, CommandError
import mpd
import sys
import pythoncom
import win32com.client
import time
import json

_mpdserver = "192.168.1.11"
_mpdport = 6600
_connectretry=3
_connectretrydelay=5
_pathfixfrom1= "D:\\Music"
_pathfixto1= "USB/music"
_pathfixfrom2= "\\"
_pathfixto2= "/"
_mpdplaylist="MusicServer Playlist"
_maxpath=256
boolReps = ['F', 'T']   # hacky!

_mmconnected = False
_mpdconnected = False
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
            MPDPlay(_sdbclient.Player.CurrentSongIndex)
        self.showMMStatus()
    def OnPause(self):      #OK
        print(">> MMEventHandlers.OnPause")
        if _sdbclient.Player.isPaused: 
            MPDPause(1)
        else:
            MPDPause(0)
        self.showMMStatus()
    def OnStop(self):
        print(">> MMEventHandlers.OnStop")
        if not _quiting:
            MPDStop()
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
        MPDSeek(_sdbclient.Player.CurrentSongIndex, int(_sdbclient.Player.PlaybackTime*.001))
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
    try:
        print(">MPDClearPlaylists()")
        mpdplaylists=MPDListPlaylists();
        for itm in mpdplaylists:
            MPDRemove(itm['playlist'])
    except (MPDError, IOError):
        print("   ! Error calling MPDClearPlaylists()") 
    print("<MPDClearPlaylists()")        
        
def SyncMMPlaylistToMPD():

    #Remove current playlist
    mpdplaylists=MPDListPlaylist();
    for itm in mpdplaylists:
        MPDRemove(itm['playlist'])
    
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
                MPDClear()
                tracks=itm.Tracks

                #Iterate through Playlist Tracks
                for k in range(0, tracks.Count):
                    itm=tracks.Item(k)
                    mpdtrack=itm.Path
                    fixedmpdtrack=FixString (mpdtrack)
                    print ("   ", fixedmpdtrack)
                    try: 
                        MPDAdd(fixedmpdtrack)
                    except: 
                        pass 
                MPDSave(playlisttitle)
            break

    
#todo: stop if current song is removed  
def SyncMMNowPlayingToMPD():
    print("?MM Count:" , _sdbclient.Player.CurrentSongList.Count)
    mpdcount=int(MPDStatus()['playlistlength'])
    print ("?MPD Count ", mpdcount)
    if _sdbclient.Player.CurrentSongList.Count == 0:
        # playlist cleared
        print ("Playlist Changed: Cleared")
        MPDClear()	
        MPDStop()
    else:
        if _sdbclient.Player.CurrentSongList.Count == mpdcount or _sdbclient.Player.CurrentSongList.Count > mpdcount:
             #songs added  or moved in playlist
            if mpdcount == 0:
                print ("Playlist Changed: MPD Playlist Empty")
                for i in range(0, _sdbclient.Player.CurrentSongList.Count):
                    mmsong=_sdbclient.Player.CurrentSongList.Item(i).Path[:_maxpath]
                    fixedmmsong=FixString (mmsong)
                    print("   + ", fixedmmsong)
                    MPDAdd(fixedmmsong)
            else:
                    print ("Playlist changed: Song added or moved")
                    for mmindex in range(0, _sdbclient.Player.CurrentSongList.Count):
                        mmsong=_sdbclient.Player.CurrentSongList.Item(mmindex).Path[:128]
                        fixedmmsong=FixString (mmsong)
                        mpdcount=int(MPDStatus()['playlistlength'])
                        found=0
                        for mpdindex in range(mmindex, mpdcount):
                            mpdsong=MPDPlaylist()[mpdindex]
                            fixedmpdsong=mpdsong.replace("file: ", "")
                            if fixedmmsong == fixedmpdsong:
                                found=1
                                if mmindex == mpdindex:
                                    print ("Ok: ", fixedmmsong)
                                    break
                                else:
                                    print ("   m", fixedmmsong)
                                    MPDMove(mpdindex, mmindex)
                        if not found:
                            MPDAdd(fixedmmsong)
                            print ("   +", fixedmmsong)
                            if mpdcount != mmindex:
                                MPDMove(mpdcount, mmindex)
                                print ("   m", fixedmmsong)
        else:
            #songs removed from playlist
            print ("Playlist Changed: Songs Removed")
            i=0
            while i < mpdcount:
                mpdsong=MPDPlaylist()[i]
                fixedmpdsong=mpdsong.replace("file: ", "")
                if i <_sdbclient.Player.CurrentSongList.Count:
                    mmsong=_sdbclient.Player.CurrentSongList.Item(i).Path[:_maxpath]
                    fixedmmsong=FixString (mmsong)
                    if fixedmmsong == fixedmpdsong:
                        print ("Ok: ", fixedmmsong)
                        i=i+1
                    else:
                        MPDDelete(i)
                        mpdcount=mpdcount-1
                        print("   -", fixedmpdsong)
                else:
                    MPDDelete(i)
                    mpdcount=mpdcount-1
                    print("   -", fixedmpdsong)
        
    # Check sync
    syncerror=False
    if _sdbclient.Player.CurrentSongList.Count != int(MPDStatus()['playlistlength']):
        syncerror=True    
    else:
        for i in range(0, _sdbclient.Player.CurrentSongList.Count):
            mpdsong=MPDPlaylist()[i]
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

def MMConnect(withevents):
    global _sdbclient
    if withevents == True:
        _sdbclient = win32com.client.DispatchWithEvents('SongsDB.SDBApplication', MMEventHandlers)
        print ("** Connected to MediaMonkey with events ***")
    else:
        _sdbclient = win32com.client.Dispatch("SongsDB.SDBApplication")
        print ("** Connected to MediaMonkey ***")
    _mmconnected=True
   
def MPDISConnect():
    try:
        _mpdclient.ping()
        #print ("?MPDISConnect()=", True)
        #print ("?MPDISConnect()=", True)
    except Exception as e:
        print ("?MPDISConnect()=", False)
        return False
    return True

def MPDConnect():
    global _mpdclient

    if MPDISConnect():
        return 0
    for i in range(0,_connectretry):
        try:
            _mpdclient = mpd.MPDClient(use_unicode=True)
            _mpdclient.connect(_mpdserver, _mpdport)
            _mpdconnected=True
            print ("** Connected to MPD server ***")
            break            
        except (MPDError, IOError):
            print("Retry:", i+1, " - Unable to connect to: ", _mpdserver,":", _mpdport)
            time.sleep(_connectretrydelay)
            pass
    if i == _connectretry-1:
        _mpdconnected=False
        sys.exit(1)
        
def MPDDisconnect():
        # Try to tell MPD we're closing the connection first
        global _mpdclient
        try:
            _mpdclient.close()
        # If that fails, don't worry, just ignore it and disconnect
        except (MPDError, IOError):
            pass

        try:
            _mpdclient.disconnect()
            print ("** Disconnected from MPD server ***")

        # Disconnecting failed, so use a new client object instead
        # This should never happen.  If it does, something is seriously broken,
        # and the client object shouldn't be trusted to be re-used.
        except (MPDError, IOError):
            print ("** Disconnected from MPD server (unexpected MPDError) ***")
            _mpdclient = MPDClient()
            pass        
        _mpdconnected=False
            

def MPDPlaylist():
    try:
        MPDConnect() #make sure we are connected to the MPD server
        return _mpdclient.playlist()
    except (MPDError, IOError):
        print ("! unhandled error in MPDPlaylist()")

def MPDListPlaylist():
    try:
        MPDConnect() #make sure we are connected to the MPD server
        return _mpdclient.listplaylists()
    except (MPDError, IOError):
        print ("! unhandled error in MPDListPlaylist()")
               
def MPDStatus():
    try:
        MPDConnect() #make sure we are connected to the MPD server
        return _mpdclient.status()
    except (MPDError, IOError):
        print ("! unhandled error in MPDStatus()")
        
def MPDRemove(item):
    try:
        MPDConnect() #make sure we are connected to the MPD server
        _mpdclient.rm(item)
    except (MPDError, IOError):
        print ("! unhandled error in MPDRemove()")
        
def MPDPlay(index):
    try:
        MPDConnect() #make sure we are connected to the MPD server
        _mpdclient.play(index)
    except (MPDError, IOError):
        print ("! unhandled error in MPDPlay()")

def MPDSeek(songindex, val):
    try:
        MPDConnect() #make sure we are connected to the MPD server
        _mpdclient.seek(songindex, val)
    except (MPDError, IOError):
        print ("! unhandled error in MPDSeek()")

def MPDPause(val):
    try:
        MPDConnect() #make sure we are connected to the MPD server
        _mpdclient.pause(val)
    except (MPDError, IOError):
        print ("! unhandled error in MPDPlay()")
        
def MPDStop():
    try:
        MPDConnect() #make sure we are connected to the MPD server
        _mpdclient.stop()
    except (MPDError, IOError):
        print ("! unhandled error in MPDStop()")

def MPDClear():
    try:
        MPDConnect() #make sure we are connected to the MPD server
        _mpdclient.clear()
    except (MPDError, IOError):
        print ("! unhandled error in MPDClear()")
        
def MPDDelete(track):
    try:
        MPDConnect() #make sure we are connected to the MPD server
        _mpdclient.delete(track)
    except (MPDError, IOError):
        print ("! unhandled error in MPDDelete()")

def MPDAdd(track):
    try:
        MPDConnect() #make sure we are connected to the MPD server
        _mpdclient.add(track)
    except (MPDError, IOError):
        print ("! unhandled error in MPDAdd()")

def MPDMove(oldindex,newindex):
    try:
        MPDConnect() #make sure we are connected to the MPD server
        _mpdclient.move(oldindex, newindex)
    except (MPDError, IOError):
        print ("! unhandled error in MPDMove()")

def MPDSave(playlisttitle):
    try:
        MPDConnect() #make sure we are connected to the MPD server
        _mpdclient.save(playlisttitle)
    except (MPDError, IOError):
        print ("! unhandled error in MPDSave)")
        
def StartMMMonitor():
    # note: once started, script does not exit until MM is shut down.
    # running the script will start MM if it's not already running
    try:
        #Connect to MPD Server and MM
        MPDConnect()
        MMConnect(True)
        
        #Stop playig and sync MPD playlist 
        _mpdclient.stop()
        _mpdclient.clear()
        SyncMMNowPlayingToMPD()

        while not _quiting:
            # required by this script because no other message loop running
            # if the app has its message loop (i.e., has a Windows UI), then
            # the events will arrive with no additional handling
            pythoncom.PumpWaitingMessages()
            time.sleep(0.2)

    except (MPDError, IOError):
        #if there is a error try to restart monitor
        StartMMMonitor()
 
    # note that SDB instance includes members of of the MMEventHandlers class
    print ("** monitor stopped; received " + str(_sdbclient._play_events) + " play events ***")
 
def Main():
   
    #handle command line argument
    total = len(sys.argv)
    cmdargs = str(sys.argv)

    #default action, no comand line argument
    if total == 1:   
        StartMMMonitor()
        #MMConnect(False)
        #time.sleep(5)
        #SyncMMNowPlayingToMPD()
        #SyncMMPlaylistToMPD()
        #print (_mpdclient.stats())

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
            elif str(sys.argv[i]) == '-stats':
                _mpdclient.stats();
            elif str(sys.argv[i]) == '-next':
                _mpdclient.next();
            elif str(sys.argv[i]) == '-previous':
                _mpdclient.previous();
            elif str(sys.argv[i]) == '-syncplaylists':
                MMConnect(False)
                SyncMMPlaylistToMPD()
            elif str(sys.argv[i]) == '-syncnowplaying':
                MMConnect(False)
                SyncMMNowPlayingToMPD()
    MPDDisconnect()
    sys.exit(0)

if __name__ == '__main__':
        
    try:
        Main()
    except Exception as e:
        print("***** Unexpected exception: %s" % e, file=sys.stderr)
        sys.exit(1)
        