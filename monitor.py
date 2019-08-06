import os

import win32file
import win32con
import subprocess
ACTIONS = {
  1 : "Created",
  2 : "Deleted",
  3 : "Updated",
  4 : "Renamed from something",
  5 : "Renamed to something"
}

FILE_LIST_DIRECTORY = 0x0001

path_to_watch = r"C:\Users\manoranjann\Documents\to_parse"
hDir = win32file.CreateFile (
  path_to_watch,
  FILE_LIST_DIRECTORY,
   win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
  # win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE ,
  None,
  win32con.OPEN_EXISTING,
  win32con.FILE_FLAG_BACKUP_SEMANTICS,
  None
)
while 1:

  results = win32file.ReadDirectoryChangesW (
    hDir,
    1024,
    True,
    win32con.FILE_NOTIFY_CHANGE_FILE_NAME |
     win32con.FILE_NOTIFY_CHANGE_DIR_NAME |
     win32con.FILE_NOTIFY_CHANGE_ATTRIBUTES |
     win32con.FILE_NOTIFY_CHANGE_SIZE |
     win32con.FILE_NOTIFY_CHANGE_LAST_WRITE |
     win32con.FILE_NOTIFY_CHANGE_SECURITY,
    None,
    None
  )
  print(results)
  for action, file in results:
    full_filename = os.path.join (path_to_watch, file)
    print(full_filename, ACTIONS.get (action, "Unknown"))
    if full_filename.endswith('.txt') and action==1:
     subprocess.call(["python", "parsing_txt1.py"])
#     subprocess.call([r'C:\Users\manoranjann\Documents\to_parse\run.bat'])
     full_filename=' '
