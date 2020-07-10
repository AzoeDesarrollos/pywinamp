#    Copyright (C) 2009 Yaron Inger, http://ingeration.blogspot.com
#
#    This program is free software; you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation; either version 2 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program; if not, write to the Free Software
#    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
#    or download it from http://www.gnu.org/licenses/gpl.txt

#    Revision History:
#
#    11/05/2008 Version 0.1
#    	- Initial release
#
#    31/12/2009 Version 0.2
#    	- Added support for keyword queries (queryAsKeyword)
#
#    10/7/2020 Version 0.3
#       - Updated code to Python 3.x and corrected PEP8 violations
#       - query() and enqueue_file() methods don't seem to work.

from ctypes import *
import win32api
import win32con
import win32gui
import win32process


class Winamp(object):
    # Winamp main window IPC
    WM_WA_IPC = win32con.WM_USER
    # Winamp's media library IPC
    WM_ML_IPC = WM_WA_IPC + 0x1000

    # add item to playlist
    IPC_ENQUEUEFILE = 100
    # playback status
    IPC_ISPLAYING = 104
    # delete winamp's internal playlist
    IPC_DELETE = 101
    # get output time
    IPC_GETOUTPUTTIME = 105
    # sets playlist position
    IPC_SETPLAYLISTPOS = 121
    # sets volume
    IPC_SETVOLUME = 122
    # gets list length
    IPC_GETLISTLENGTH = 124
    # gets given index playlist file
    IPC_GETPLAYLISTFILE = 211
    # gets given index playlist title
    IPC_GETPLAYLISTTITLE = 212
    # next playlist track
    IPC_GET_NEXT_PLITEM = 361
    # current playing title
    IPC_GET_PLAYING_TITLE = 3034

    # runs an ml query
    ML_IPC_DB_RUNQUERY = 0x0700
    # runs an ml query as a string
    ML_IPC_DB_RUNQUERY_SEARCH = 0x0701
    # frees ml query results from winamp's memory
    ML_IPC_DB_FREEQUERYRESULTS = 0x0705

    # playback possible options
    PLAYBACK_NOT_PLAYING = 0
    PLAYBACK_PLAYING = 1
    PLAYBACK_PAUSE = 3

    GET_TRACK_POSITION = 0
    GET_TRACK_LENGTH = 1

    # winamp buttons commands
    BUTTON_COMMAND_PREVIOUS = 40044
    BUTTON_COMMAND_PLAY = 40045
    BUTTON_COMMAND_PAUSE = 40046
    BUTTON_COMMAND_STOP = 40047
    BUTTON_COMMAND_NEXT = 40048

    # sort playlist by path and filename
    ID_PE_S_PATH = 40211

    class COPYDATATYPE(Structure):
        _fields_ = [("szData", c_char_p)]

    """
    typedef struct tagCOPYDATASTRUCT {
        ULONG_PTR dwData;
        DWORD cbData;
        PVOID lpData;
    } COPYDATASTRUCT, *PCOPYDATASTRUCT;

    dwData
        Specifies data to be passed to the receiving application. 
    cbData
        Specifies the size, in bytes, of the data pointed to by the lpData member. 
    lpData
        Pointer to data to be passed to the receiving application. This member can be NULL.
    """

    class COPYDATASTRUCT(Structure):
        _fields_ = [("dwData", c_ulong),
                    ("cbData", c_ulong),
                    ("lpData", c_void_p)]

    """
    typedef struct 
    {
      itemRecord *Items;
      int Size;
      int Alloc;
    } itemRecordList;
    """

    class ItemRecordList(Structure):
        _fields_ = [("Items", c_void_p),
                    ("Size", c_int),
                    ("Alloc", c_int)]

    """
    typedef struct
    {
      char *filename;
      char *title;
      char *album;
      char *artist;
      char *comment;
      char *genre;
      int year;
      int track;
      int length;
      char **extended_info;
      // currently defined extended columns (while they are stored internally as integers
      // they are passed using extended_info as strings):
      // use getRecordExtendedItem and setRecordExtendedItem to get/set.
      // for your own internal use, you can set other things, but the following values
      // are what we use at the moment. Note that setting other things will be ignored
      // by ML_IPC_DB*.
      // 
      //"RATING" file rating. can be 1-5, or 0 or empty for undefined
      //"PLAYCOUNT" number of file plays.
      //"LASTPLAY" last time played, in standard time_t format
      //"LASTUPD" last time updated in library, in standard time_t format
      //"FILETIME" last known file time of file, in standard time_t format
      //"FILESIZE" last known file size, in kilobytes.
      //"BITRATE" file bitrate, in kbps
        //"TYPE" - "0" for audio, "1" for video

    } itemRecord;
    """

    class ItemRecord(Structure):
        _fields_ = [("filename", c_char_p),
                    ("title", c_char_p),
                    ("album", c_char_p),
                    ("artist", c_char_p),
                    ("comment", c_char_p),
                    ("genre", c_char_p),
                    ("year", c_int),
                    ("track", c_int),
                    ("length", c_int),
                    ("extended_info", c_char_p)]

    """
    typedef struct 
    {
      char *query;
      int max_results;      // can be 0 for unlimited
      itemRecordList results;
    } mlQueryStruct;
    """

    class MlQueryStruct(Structure):
        pass

    playlist = None

    def __init__(self):
        self.__init_structures()

        # get important Winamp's window handles
        try:
            self.__mainWindowHWND = self.__find_window([("Winamp v1.x", None)])
            self.__playlistHWND = self.__find_window([("BaseWindow_RootWnd", None),
                                                      ("BaseWindow_RootWnd", "Playlist Editor"),
                                                      ("Winamp PE", "Winamp Playlist Editor")])
            self.__mediaLibraryHWND = self.__find_window([("BaseWindow_RootWnd", None),
                                                          ("BaseWindow_RootWnd", "Winamp Library"),
                                                          ("Winamp Gen", "Winamp Library"), (None, None)])
        except RuntimeError:
            raise Exception("Cannot find Winamp windows. Is winamp started?")

        self.__processID = win32process.GetWindowThreadProcessId(self.__mainWindowHWND)[1]

        # open Winamp's process
        self.__hProcess = windll.kernel32.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, self.__processID)

    def detach(self):
        """Detaches from Winamp's process."""
        windll.kernel32.CloseHandle(self.__hProcess)

    def __init_structures(self):
        # cannot be done statically, because mlQueryStruct doesn't know any self.itemRecordList
        try:
            self.MlQueryStruct._fields_ = [("query", c_char_p),
                                           ("max_results", c_int),
                                           ("ItemRecordList", self.ItemRecordList)]
        except AttributeError:
            # mlQueryStruct already initialized
            pass

    @staticmethod
    def __find_window(window_list):
        """Gets a handle to the lowest window in the given windows hierarchy.

        The given list should be in format of [<Window Class>, <Window Name>]."""

        current_window = None

        for i in range(len(window_list)):
            if current_window is None:
                current_window = win32gui.FindWindow(window_list[i][0], window_list[i][1])
            else:
                current_window = win32gui.FindWindowEx(current_window, 0, window_list[i][0], window_list[i][1])

        return current_window

    def __setattr__(self, attr, value):
        if attr == "playlist":
            self.clear_playlist()

            [self.enqueue_file(item.filename) for item in value]
        else:
            object.__setattr__(self, attr, value)

    def __getattr__(self, attr):
        print(attr)
        if attr == "playlist":
            return self.get_playlist_filenames()
        else:
            try:
                return getattr(self, attr)
            except Exception:
                raise AttributeError(attr)

    def __read_string_from_memory(self, address, is_unicode=False):
        """Reads a string from Winamp's memory address space."""

        if is_unicode:
            buffer_length = win32con.MAX_PATH * 2
            buffer = create_unicode_buffer(buffer_length * 2)
        else:
            buffer_length = win32con.MAX_PATH
            buffer = create_string_buffer(buffer_length)

        bytes_read = c_ulong(0)

        """Note: this is quite an ugly hack, because we assume the string will have a maximum 
        size of MAX_PATH (=260) and that we're not in an end of a page.
        
        A proper way to solve it would be:
            1. calling VirutalQuery.
            2. Reading one byte at a time. (?)
            3. Use CreateRemoteThread to run strlen on Winamp's process."""

        windll.kernel32.ReadProcessMemory(self.__hProcess, address, buffer, buffer_length, byref(bytes_read))

        return buffer.value

    def enqueue_file(self, file_path):
        """Enqueues a file in Winamp's playlist.

        filePath is the given file path to enqueue."""

        # prepare copydata structure for sending data
        cpy_data = create_string_buffer(file_path)
        cds = self.COPYDATASTRUCT(c_ulong(self.IPC_ENQUEUEFILE),
                                  c_ulong(len(cpy_data.raw)),
                                  cast(cpy_data, c_void_p))

        # send copydata message
        win32api.SendMessage(self.__mainWindowHWND, win32con.WM_COPYDATA, 0, addressof(cds))

    def query(self, query_string, query_type=ML_IPC_DB_RUNQUERY):
        """Queries Winamp's media library and returns a list of items matching the query.

        The query should include filters like '?artist has \'alice\''.
        For more information, consult your local Winamp forums or media library."""

        query_string_addr = self.__copy_data_to_winamp(query_string)

        # create query structs and copy to winamp
        record_list = self.ItemRecordList(0, 0, 0)
        query_struct = self.MlQueryStruct(cast(query_string_addr, c_char_p), 0, record_list)
        query_struct_addr = self.__copy_data_to_winamp(query_struct)

        # run query
        win32api.SendMessage(self.__mediaLibraryHWND, self.WM_ML_IPC, query_struct_addr, query_type)

        received_query = self.__read_data_from_winamp(query_struct_addr, self.MlQueryStruct)

        items = []

        buf = create_string_buffer(sizeof(self.itemRecord) * received_query.itemRecordList.Size)
        windll.kernel32.ReadProcessMemory(self.__hProcess, received_query.itemRecordList.Items, buf, sizeof(buf), 0)

        for i in range(received_query.itemRecordList.Size):
            item = self.__read_data_from_winamp(received_query.itemRecordList.Items + (sizeof(self.itemRecord) * i),
                                                self.itemRecord)

            self.__fix_remote_struct(item)

            items.append(item)

        # free results
        win32api.SendMessage(self.__mediaLibraryHWND, self.WM_ML_IPC, query_struct_addr,
                             self.ML_IPC_DB_FREEQUERYRESULTS)

        return items

    def query_as_keyword(self, query_string):
        """Queries Winamp's media library and returns a list of items matching the query.

        The query should be a keyword, like 'alice' (then the query is then treated as a string query).
        This makes Winamp search the requested keyword in every data field in the media library database
        (such as Artist, Album, Track Name, ...)."""

        return self.query(query_string, self.ML_IPC_DB_RUNQUERY_SEARCH)

    def __copy_data_to_winamp(self, data):
        if type(data) is str:
            data_to_copy = create_string_buffer(bytes(data))
        else:
            data_to_copy = data

        # allocate data in Winamp's address space
        lp_address = windll.kernel32.VirtualAllocEx(self.__hProcess, None, sizeof(data_to_copy), win32con.MEM_COMMIT,
                                                    win32con.PAGE_READWRITE)
        # write data to Winamp's memory
        windll.kernel32.WriteProcessMemory(self.__hProcess, lp_address, addressof(data_to_copy), sizeof(data_to_copy),
                                           None)

        return lp_address

    def __read_data_from_winamp(self, address, ctypes_type):
        if ctypes_type is c_char_p:
            buffer = create_string_buffer(win32con.MAX_PATH)
        else:
            buffer = create_string_buffer(sizeof(ctypes_type))

        bytes_read = c_ulong(0)
        if windll.kernel32.ReadProcessMemory(self.__hProcess, address, buffer, sizeof(buffer), byref(bytes_read)) == 0:
            # we're in a new page
            if address / 0x1000 != address + sizeof(buffer):
                # possible got into an unpaged memory region, read until end of page
                windll.kernel32.ReadProcessMemory(self.__hProcess, address, buffer,
                                                  ((address + 0x1000) & 0xfffff000) - address, byref(bytes_read))
            else:
                raise RuntimeError("ReadProcessMemory failed at 0x%08x." % address)

        if ctypes_type is c_char_p:
            return cast(buffer, c_char_p)
        else:
            return cast(buffer, POINTER(ctypes_type))[0]

    def __fix_remote_struct(self, structure):
        offset = 0

        for i in range(len(structure._fields_)):
            (field_name, field_type) = structure._fields_[i]

            if field_type is c_char_p or field_type is c_void_p:
                # get pointer address
                address = cast(addressof(structure) + offset, POINTER(c_int))[0]

                # ignore null pointers
                if address != 0x0:
                    structure.__setattr__(field_name, self.__read_data_from_winamp(address, field_type))

            offset += sizeof(field_type)

    def __send_user_message(self, w_param, l_param, hwnd=None):
        """Sends a user message to the given hwnd with the given wParam and lParam."""
        if hwnd is None:
            target_hwnd = self.__mainWindowHWND
        else:
            target_hwnd = hwnd

        return win32api.SendMessage(target_hwnd, self.WM_WA_IPC, w_param, l_param)

    def __send_command_message(self, w_param, l_param, hwnd=None):
        """Sends a command message to the given hwnd with the given wParam and lParam."""
        if hwnd is None:
            target_hwnd = self.__mainWindowHWND
        else:
            target_hwnd = hwnd

        return win32api.SendMessage(target_hwnd, win32con.WM_COMMAND, w_param, l_param)

    def get_playback_status(self):
        """Gets Winamp's playback status. Use the constants PLAYBACK_NOT_PLAYING,
        PLAYBACK_PLAYING and PLAYBACK_PAUSE = 3 to resolve status."""
        return self.__send_user_message(0, self.IPC_ISPLAYING)

    def get_playing_track_length(self):
        """Gets the length in second of the current playing track."""
        return self.__send_user_message(self.GET_TRACK_LENGTH, self.IPC_GETOUTPUTTIME)

    def get_playing_track_position(self):
        """Gets the position in milliseconds of the current playing track."""
        return self.__send_user_message(self.GET_TRACK_POSITION, self.IPC_GETOUTPUTTIME)

    def clear_playlist(self):
        """Clears the playlist."""
        return self.__send_user_message(0, self.IPC_DELETE)

    def set_playlist_position(self, position):
        """Sets the playlist position in the given position number (zero based)."""
        return self.__send_user_message(position, self.IPC_SETPLAYLISTPOS)

    def set_volume(self, volume):
        """Sets the internal Winamp's volume meter. Will only accept values in the range 0-255."""
        assert 0 <= volume <= 255

        self.__send_user_message(volume, self.IPC_SETVOLUME)

    def get_current_playing_title(self):
        """Returns the title of the current playing track"""
        address = self.__send_user_message(0, self.IPC_GET_PLAYING_TITLE)

        return self.__read_string_from_memory(address, True)

    def get_playlist_file(self, position):
        """Returns the filename of the current selected file in the playlist."""
        address = self.__send_user_message(position, self.IPC_GETPLAYLISTFILE)

        return self.__read_string_from_memory(address)

    def get_playlist_title(self, position):
        """Returns the title of the current selected file in the playlist."""
        address = self.__send_user_message(position, self.IPC_GETPLAYLISTTITLE)

        return self.__read_string_from_memory(address)

    def get_list_length(self):
        """Returns the length of the current playlist."""
        return self.__send_user_message(0, self.IPC_GETLISTLENGTH)

    def get_playlist_filenames(self):
        """Retrieves a list of the current playlist song filenames."""
        return [self.get_playlist_file(position) for position in range(self.get_list_length())]

    def get_playlist_titles(self):
        """Retrieves a list of the current playlist song titles."""
        return [self.get_playlist_title(position) for position in range(self.get_list_length())]

    def next(self):
        """Sets playlist marker to next song."""
        self.__send_command_message(self.BUTTON_COMMAND_NEXT, 0)

    def previous(self):
        """Sets playlist marker to previous song."""
        self.__send_command_message(self.BUTTON_COMMAND_PREVIOUS, 0)

    def pause(self):
        """Pauses Winamp's playback."""
        self.__send_command_message(self.BUTTON_COMMAND_PAUSE, 0)

    def play(self):
        """Starts / resumes playing Winamp's playback, or restarts current playing song."""
        self.__send_command_message(self.BUTTON_COMMAND_PLAY, 0)

    def stop(self):
        """Stops Winamp's playback."""
        self.__send_command_message(self.BUTTON_COMMAND_STOP, 0)

    def sort_playlist(self):
        """Sorts the current playlist alphabetically."""
        self.__send_command_message(self.ID_PE_S_PATH, 0, self.__playlistHWND)

    def play_album(self, album):
        """Plays a given album name."""
        self.playlist = self.query("album = \"%s\"" % album)
        self.stop()
        self.sortPlaylist()
        self.setPlaylistPosition(0)
        self.play()


def print_media_library_item(item):
    print("Filename: %s\nTrack: %s, Album: %s, Artist: %s\nComment: %s, Genre: %s" % (
        item.filename, item.track, item.album, item.artist, item.comment, item.genre))


if __name__ == "__main__":
    # little demonstration...

    w = Winamp()

    # print("Current playlist:")
    # print(w.get_playlist_titles())

    # items = w.query("artist has \"opeth\"")
    # [printMediaLibraryItem(item) for item in items]

    # w.playlist = w.query("artist has \"jane's\"")
    # w.sort_playlist()

    # print(w.playlist)

    w.enqueue_file(b'D:/mp3/Air - Tory no Uta (full).mp3')

    # "Playing album..."
    # w.playAlbum("Red")
