using System;
using System.Xml.Serialization;
using System.Xml;
using System.IO;
using System.Security.Cryptography;

namespace MyRenamer
{
    public class Config
    {
        [NonSerialized()]

        #region Fields
        private static string _fileName;
        private static Config instance = new Config();
        #endregion

        #region config public properties
        public string SqlDatabase { get; set; }
        public string Caption { get; set; }
        #endregion Properties

        #region CTOR
        private Config()
        {  }
        #endregion

        #region Methods
        public static Config GetInstance(string filename)
        {
            _fileName = filename;
            // Config retVal = null;
            XmlSerializer serializer = new XmlSerializer(typeof(Config));
            using (StreamReader reader = new StreamReader(_fileName))
            { 
              try
                { 
                    instance = (Config)serializer.Deserialize(reader);
                }
                catch
                { }
            }
            return instance;
        }
        #endregion

    }
}
