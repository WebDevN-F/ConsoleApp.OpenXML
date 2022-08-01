using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace ConsoleApp.OpenXML
{
    public partial class Wordprocessing
    {
        public static void SearchAndReplace(string document, Dictionary<string, string> dict)
        {           
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                var jsonString = ParseXmlToJSON(docText);

                foreach (KeyValuePair<string, string> item in dict)
                {
                    Regex regexText = new Regex(item.Key);
                    docText = regexText.Replace(docText, item.Value);
                }

                var jsonString2 = ParseXmlToJSON(docText);

                using (StreamWriter sw = new StreamWriter(
                          wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

        public static string ParseXmlToJSON(string docText)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(docText);
            string jsonText = JsonConvert.SerializeXmlNode(xmlDoc);
            return jsonText;
        }

        public static XmlDocument DeserializeXmlNode(string jsonDoc)
        {
            XmlDocument doc = (XmlDocument)JsonConvert.DeserializeXmlNode(jsonDoc);
            return doc;
        }

        // https://products.fileformat.com/word-processing/net/docx-to-pdf-converter/
        public static void PrintToPdf(string fileDocx)
        {


        }
    }
}
