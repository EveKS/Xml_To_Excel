using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using Xml2CSharp;

namespace Xml_To_Excel.Utility
{
    public class ReadXmlFoder
    {
        async void Read()
        {
            string directory = @"D:\Test";
            foreach (var fi in new DirectoryInfo(directory)
                .EnumerateFiles("*.xml", SearchOption.AllDirectories))
            {
                #region FromXml
                XmlSerializer formatter = new XmlSerializer(typeof(Bill));
                string tmp = string.Empty;
                using (StreamReader sr = new StreamReader(fi.FullName, Encoding.Default))
                {
                    tmp += await sr.ReadToEndAsync();
                }
                Bill bill;
                XmlDocument xDoc = new XmlDocument();
                xDoc.LoadXml(tmp.Replace("\x0c", ""));
                var test = DeserializeFromXmlDocument(xDoc);
                using (FileStream fs = new FileStream(@"01.2016_New.xml", FileMode.Open))
                {
                    XmlReader reader = XmlReader.Create(xDoc.InnerText,);
                    bill = (Bill)formatter.Deserialize(reader);
                }
                #endregion
            }
        }
            static IEnumerable<Bill> DeserializeFromXmlDocument(XmlDocument doc)
        {
                XmlSerializer seri = new XmlSerializer(typeof(Bill));

                using (var reader = new XmlNodeReader(doc.DocumentElement))
                {
                    reader.MoveToContent();
                    reader.ReadStartElement();
                    while (reader.IsStartElement())
                    {
                        Bill entry = (Bill)seri.Deserialize(reader);
                        yield return entry;
                    }
                }
            }
        }
    }

