# A solution sample generate document from template office word & binding Dictionary data keyvalues

## OpenXML 

- Packages:

```
  <ItemGroup>
    <PackageReference Include="DocumentFormat.OpenXml" Version="2.17.1" />
    <PackageReference Include="Newtonsoft.Json" Version="13.0.1" />
  </ItemGroup>
```

-- LOGIC

```
    public static void SearchAndReplace(string document, Dictionary<string, string> dict)
        {           
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                // var jsonString = ParseXmlToJSON(docText);

                foreach (KeyValuePair<string, string> item in dict)
                {
                    Regex regexText = new Regex(item.Key);
                    docText = regexText.Replace(docText, item.Value);
                }

                // var jsonString2 = ParseXmlToJSON(docText);

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }
```

## Convert document docx to pdf opensource & LibreOfficePortable

https://products.fileformat.com/word-processing/net/docx-to-pdf-converter/