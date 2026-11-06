using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

public class Program
{
    public static void Main()
    {
        // Create a temporary working directory
        string workDir = Path.Combine(Path.GetTempPath(), "WordOleBatch_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDir);

        // Prepare a list of Word file paths
        List<string> wordFiles = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            string filePath = Path.Combine(workDir, $"Document{i}.docx");
            CreateWordDocument(filePath);
            wordFiles.Add(filePath);
        }

        // Create a single minimal Excel file (empty zip) in memory (identical for all documents)
        byte[] excelBytes = CreateEmptyExcelPackage();

        // Insert the Excel OLE object into each Word document
        foreach (string docPath in wordFiles)
        {
            // Open the existing .docx (zip) and add the embedded Excel part and relationships
            using (FileStream fs = new FileStream(docPath, FileMode.Open, FileAccess.ReadWrite))
            using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Update))
            {
                // Add the embedded Excel file
                string embedPath = "word/embeddings/embedded.xlsx";
                ZipArchiveEntry embedEntry = archive.GetEntry(embedPath) ?? archive.CreateEntry(embedPath);
                using (Stream es = embedEntry.Open())
                {
                    es.SetLength(0);
                    es.Write(excelBytes, 0, excelBytes.Length);
                }

                // Add /word/_rels/document.xml.rels if it does not exist
                string relsPath = "word/_rels/document.xml.rels";
                ZipArchiveEntry relsEntry = archive.GetEntry(relsPath) ?? archive.CreateEntry(relsPath);
                using (Stream rs = relsEntry.Open())
                using (StreamWriter writer = new StreamWriter(rs))
                {
                    rs.SetLength(0);
                    writer.Write(GetDocumentRelsXml());
                }

                // Update /word/document.xml to reference the OLE object
                string docXmlPath = "word/document.xml";
                ZipArchiveEntry docXmlEntry = archive.GetEntry(docXmlPath) ?? archive.CreateEntry(docXmlPath);
                using (Stream ds = docXmlEntry.Open())
                using (StreamWriter writer = new StreamWriter(ds))
                {
                    ds.SetLength(0);
                    writer.Write(GetDocumentXml());
                }
            }
        }

        // Clean up: delete the temporary directory and its contents
        // (Comment out the following line if you want to inspect the generated files)
        Directory.Delete(workDir, true);
    }

    // Creates a minimal .docx file with the required parts but without the embedded object.
    private static void CreateWordDocument(string filePath)
    {
        using (FileStream fs = new FileStream(filePath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Create))
        {
            // [Content_Types].xml
            var ctEntry = archive.CreateEntry("[Content_Types].xml");
            using (var writer = new StreamWriter(ctEntry.Open()))
            {
                writer.Write(GetContentTypesXml());
            }

            // _rels/.rels
            var relsEntry = archive.CreateEntry("_rels/.rels");
            using (var writer = new StreamWriter(relsEntry.Open()))
            {
                writer.Write(GetPackageRelsXml());
            }

            // word/document.xml (placeholder without OLE)
            var docEntry = archive.CreateEntry("word/document.xml");
            using (var writer = new StreamWriter(docEntry.Open()))
            {
                writer.Write(GetDocumentXml());
            }

            // word/_rels/document.xml.rels (placeholder)
            var docRelsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
            using (var writer = new StreamWriter(docRelsEntry.Open()))
            {
                writer.Write(GetDocumentRelsXml());
            }
        }
    }

    // Returns a minimal empty Excel file (just the ZIP EOCD record)
    private static byte[] CreateEmptyExcelPackage()
    {
        // An empty .xlsx is essentially an empty zip file. This byte array represents a minimal zip.
        return new byte[] { 0x50, 0x4B, 0x05, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
    }

    // XML for [Content_Types].xml
    private static string GetContentTypesXml()
    {
        return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/word/document.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml""/>
  <Override PartName=""/word/embeddings/embedded.xlsx"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet""/>
</Types>";
    }

    // XML for _rels/.rels
    private static string GetPackageRelsXml()
    {
        return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""word/document.xml""/>
</Relationships>";
    }

    // XML for word/document.xml (includes a placeholder for the OLE object)
    private static string GetDocumentXml()
    {
        return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
            xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
            xmlns:v=""urn:schemas-microsoft-com:vml"">
  <w:body>
    <w:p><w:r><w:t>Embedded Excel OLE Object:</w:t></w:r></w:p>
    <w:p>
      <w:r>
        <w:object>
          <v:shape id=""obj1"" type=""#_x0000_t75"" style=""width:100pt;height:100pt"">
            <v:imagedata r:id=""rId1""/>
          </v:shape>
        </w:object>
      </w:r>
    </w:p>
    <w:p><w:r><w:t>End of embedded object.</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>";
    }

    // XML for word/_rels/document.xml.rels (relationship to the embedded Excel file)
    private static string GetDocumentRelsXml()
    {
        return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"" Target=""embeddings/embedded.xlsx""/>
</Relationships>";
    }
}
