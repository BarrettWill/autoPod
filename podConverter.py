import os
import yt_dlp
import win32com.client

# download directory: C:/yt_downloads
# itunes will open up after all downloads have processed, do not be alarmed
# for multi-select inside itunes, ctrl + click on a file and it will stay selected, right click file and send to desired device


URLS = [
    'https://www.youtube.com/watch?v=KdLSSTssujo&list=RDKdLSSTssujo&start_radio=1&pp=ygUIb3NhbWFzb26gBwE%3D',                                               # single video
    'https://www.youtube.com/watch?v=MeiZOKNXsCc&list=PLXFKdhVuenFE&pp=sAgC',                                                                                      # playlist
    'https://soundcloud.com/him-124623755/californiacrazy-prod-ok-smd?utm_source=clipboard&utm_medium=text&utm_campaign=social_sharing',                    # single track
    # 'https://soundcloud.com/artist/sets/set-name',                                                                                                        # set (soundcloud playlist)
]

save_dir = r'C:\yt_downloads'
os.makedirs(save_dir, exist_ok=True)

files = []
def done(d):
    if d['status'] == 'finished':
        files.append(os.path.splitext(d['filename'])[0] + '.mp3')

def is_mix(url):
    return 'list=RD' in url   # YouTube auto-generated Mix: treat as one video

def make_opts(url):
    return {
        'format': 'bestaudio/best', # automatically finds best audio
        'noplaylist': is_mix(url), # prevents single songs from downloading everything in sight!!
        'ignoreerrors': True,      # one error from a track does not stop rest from downloading
        'download_archive': os.path.join(save_dir, 'archive.txt'),  # skip anything already downloaded
        'paths': {'home': save_dir},
        'outtmpl': '%(title)s.%(ext)s', # what the file name is saved as
        'cookiefile': os.path.join(save_dir, 'cookies.txt'), # cookies to prevent youtube from thinking you're a bot
        'remote_components': ['ejs:github'],
        'progress_hooks': [done],
        'postprocessors': [
            {'key': 'FFmpegExtractAudio', 'preferredcodec': 'mp3', 'preferredquality': '192'}, # my preferred file specifics, ipod accepts mp3
            {'key': 'FFmpegMetadata'},
        ],
    }

# go through array of mix of urls, download each url with specific attributes found in make_opts
for url in URLS:
    with yt_dlp.YoutubeDL(make_opts(url)) as ydl:
        ydl.download([url])

# push new files into iTunes and sync

# add to files array if path exists
files = [f for f in files if os.path.exists(f)]

# check if files contains any songs
if files:

    # open itunes
    itunes = win32com.client.Dispatch("iTunes.Application")
    playlist = None

    # search every playlist inside of itunes for specific playlist "autoPod"
    for pl in itunes.LibrarySource.Playlists:

        # if autoPod playlist already exists
        if pl.Name == "autoPod":
            playlist = pl
            break

        # if autopod playlist does not already exist, create it
    if playlist is None:
        playlist = itunes.CreatePlaylist("autoPod")

    for f in files:
        playlist.AddFile(os.path.abspath(f))

    # sync ipod, some ipod versions do not have auto sync (mine does not(ipod shuffle))
    itunes.UpdateIPod()
    print("Added", len(files), "song(s) and started sync.")

# no new songs
else:
    print("Nothing new to add.")
