using System;
using System.IO;
using System.Xml.Serialization;
using System.Xml;

namespace OutlookGoogleSync
{
    /// <summary>
    /// Exports or imports any object to/from XML.
    /// </summary>
    public class XmlManager
    {       
        /// <summary>
        /// Exports any object given in "obj" to an xml file given in "filename"
        /// </summary>
        /// <param name="obj">The object that is to be serialized/exported to XML.</param>
        /// <param name="filename">The filename of the xml file to be written.</param>
        public static void Export(Object obj, string filename)
        {
            using (var writer = new XmlTextWriter(filename, null))
            {
                writer.Indentation = 4;
                writer.Formatting = Formatting.Indented;
                new XmlSerializer(obj.GetType()).Serialize(writer, obj);
            }
        }
        
        /// <summary>
        /// Imports from XML and returns the resulting object of type T.
        /// </summary>
        /// <param name="filename">The XML file from which to import.</param>
        /// <returns></returns>
        public static T Import<T>(string filename)
        {
            using (var fs = new FileStream(filename, FileMode.Open))
            {
                var xmlSerializer = new XmlSerializer(typeof(T));
                var result = (T)xmlSerializer.Deserialize(fs);
                return result;
            }
        }
    }
}
