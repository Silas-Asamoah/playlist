#===== SCRIPT GENESIS===========

from bs4 import BeautifulSoup 

desired_artists = ['Imagine Dragons', 'One Republic', 'Bastille', 'Marshmello','Owl City',
                   'Kygo', 'Drake', 'J Cole', 'Hillsong Young & Free',
                   'ColdPlay', 'Logic', 'Stormzy', 'Dua Lipa',
                   'Khalid', 'Stormzy', 'Jon Bellion', 'The Chainsmokers']

desired_artists.sort()

try:
    songRequest = requests.get('')
    topSongsHtml = songRequest.text
except:
    print('error getting top 50 song list html')
    
# parsing with BeautifulSoup
html_tracks =  BeautifulSoup(topSongsHtml, 'html.parser')
        