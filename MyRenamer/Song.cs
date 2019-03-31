using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace MyRenamer
{
    class Song
    {
        string _fullpath = null;
        string _path = null;
        string _name = null;
        Guid _id;

        #region Properties
        public string FullPath
        {
            get { return _fullpath; }
        }

        //If someone need a path (perhaps the engine)
        public string Path
        {
            get { return _path; }
        }

        //If someone needs a name (perhaps the UI)
        public string Name
        {
            get { return _name; }
        }

        //If someone needs the unique ID of song
        public Guid ID
        {
            get { return _id; }
        }

        #endregion

        public Song(string fullpath)
        {
            _fullpath = fullpath;
            _name = GetNameFromPath(fullpath);
            _path = System.IO.Path.GetPathRoot(fullpath);
            _id = Guid.NewGuid();
        }

        public Song(string path, string name)
        {
            _path = path;
            _name = name;
            _fullpath = System.IO.Path.Combine(path, name);
            _id = Guid.NewGuid();
        }

        //This will convert the path to filename(windows)
        string GetNameFromPath(string path)
        {
            string name = "";
            string[] split = path.Split('\\');
            string nameExt = split[split.Length - 1];
            split = nameExt.Split('.');
            for (int i = 0; i < split.Length - 1; ++i)
            {
                name += split[i];
            }
            return name;
        }

    }
}
