using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace MyRenamer
{
    class PlayList
    {
        List<Song> playList = null;
        int currentItemIndex = -1;

        public PlayList()
        {
            playList = new List<Song>();
        }

        //Add item to playlist
        public string AddItem(string path, out string uid)
        {
            if (currentItemIndex == -1)
            {
                currentItemIndex = 0;
            }            
            Song song = new Song(path);
            playList.Add(song);
            uid = song.ID.ToString();
            return song.Name;
        }

        public string AddItem(Song song, out string uid)
        {
            if (currentItemIndex == -1)
            {
                currentItemIndex = 0;
            }
            playList.Add(song);
            uid = song.ID.ToString();
            return song.Name;
        }

        //Get the Item from playlist
        public Song GetItem(string uid)
        {
            Guid id = new Guid(uid);

            foreach (Song s in playList)
            {
                if (s.ID == id)
                {
                    return s;
                }
            }

            return null;
        }

        //Get the Item from playlist
        public Song GetNextItem()
        {
            Song s = null;
            if (currentItemIndex + 1 >= playList.Count)
            {
                return s;
            }
            else if (currentItemIndex != -1 && (currentItemIndex) != playList.Count)
            {
                ++currentItemIndex;
                s = playList[currentItemIndex];                
            }          

            return s;
        }

        //Get the Item from playlist
        public void SetCurrentItem(string uid)
        {
            int index = -1;
            foreach (Song s in playList)
            {
                ++index;
                if (s.ID.ToString() == uid)
                {
                    currentItemIndex = index;                    
                    break;
                }
            }
        }

        //Get the Item from playlist
        public Song GetPreviousItem()
        {
            Song s = null;
            if (currentItemIndex - 1 <= -1)
            {
                return s;
            }
            else if ((currentItemIndex) > -1 && currentItemIndex != playList.Count)
            {
                --currentItemIndex;
                s = playList[currentItemIndex];               
            }            

            return s;
        }

        //Get the Item from playlist
        public Song GetCurrentItem()
        {
            if (currentItemIndex != -1 && currentItemIndex != playList.Count)
            {
                return playList[currentItemIndex];
            }

            return null;
        }


        //Remove item from playlist
        public void RemoveItem(string uid)
        {
            Guid id = new Guid(uid);
            int index = -1;
            foreach (Song s in playList)
            {
                ++index;
                if (s.ID == id)
                {
                    playList.Remove(s);

                    if (index < currentItemIndex)
                    {
                        --currentItemIndex;
                    }
                    break;
                }
            }
        }

        //Clear the playlist
        public void Clear()
        {
            playList.Clear();
            currentItemIndex = -1;
        }
    }
}
