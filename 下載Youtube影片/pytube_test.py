from pytube import YouTube
from pytube import Playlist

p = Playlist('https://www.youtube.com/playlist?list=PL7QimmJ9Zfe-hn-jZxrkNmG3DX_XgLqpA')
for url in p.video_urls:
    try:
        yt = YouTube(url)
        print(f'下載中 影片: {yt.title}')
        yt.streams.get_highest_resolution().download('./videos') # 這是最高畫質，如果要特定畫質，就要先print(yt.streams)看看，然後用yt.streams.filter(res='')進一步搜尋
    except:
        print(f'影片: {yt.title} 無法下載，跳過繼續下載下一部。', end='\n\n')
    else:
        print("影片下載完成", end='\n\n')