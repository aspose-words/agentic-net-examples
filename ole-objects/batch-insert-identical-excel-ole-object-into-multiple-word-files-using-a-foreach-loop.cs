using System;
using System.IO;
using System.IO.Compression;
using System.Linq;

public class Program
{
    public static void Main()
    {
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchDemo");
        string wordDir = Path.Combine(baseDir, "WordDocs");
        string excelPath = Path.Combine(baseDir, "Sample.xlsx");

        // Prepare directories
        Directory.CreateDirectory(wordDir);

        // Create a minimal Excel file if it does not exist
        if (!File.Exists(excelPath))
        {
            CreateMinimalExcel(excelPath);
        }

        // Create sample Word documents if none exist
        if (!Directory.EnumerateFiles(wordDir, "*.docx").Any())
        {
            for (int i = 1; i <= 3; i++)
            {
                string docPath = Path.Combine(wordDir, $"Document{i}.docx");
                CreateMinimalWord(docPath, $"This is document {i}.");
            }
        }

        // Batch "embed" the Excel file into each Word document
        foreach (string docFile in Directory.EnumerateFiles(wordDir, "*.docx"))
        {
            // Open the .docx (which is a zip package) for update
            using (ZipArchive zip = ZipFile.Open(docFile, ZipArchiveMode.Update))
            {
                // 1. Add the Excel file under word/embeddings/
                string embedEntryPath = $"word/embeddings/{Path.GetFileName(excelPath)}";
                ZipArchiveEntry existingEmbed = zip.GetEntry(embedEntryPath);
                existingEmbed?.Delete();

                ZipArchiveEntry embedEntry = zip.CreateEntry(embedEntryPath, CompressionLevel.NoCompression);
                using (Stream embedStream = embedEntry.Open())
                using (FileStream src = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
                {
                    src.CopyTo(embedStream);
                }

                // 2. Insert a simple placeholder paragraph into document.xml
                const string documentXmlPath = "word/document.xml";
                ZipArchiveEntry docXmlEntry = zip.GetEntry(documentXmlPath);
                if (docXmlEntry != null)
                {
                    string xmlContent;
                    using (StreamReader reader = new StreamReader(docXmlEntry.Open()))
                    {
                        xmlContent = reader.ReadToEnd();
                    }

                    // Simple paragraph that indicates an embedded Excel object
                    string placeholderParagraph =
                        "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:r><w:t>[Embedded Excel]</w:t></w:r></w:p>";

                    int bodyCloseIndex = xmlContent.IndexOf("</w:body>", StringComparison.Ordinal);
                    if (bodyCloseIndex >= 0)
                    {
                        xmlContent = xmlContent.Insert(bodyCloseIndex, placeholderParagraph);

                        // Replace the old document.xml entry
                        docXmlEntry.Delete();
                        ZipArchiveEntry newDocXmlEntry = zip.CreateEntry(documentXmlPath, CompressionLevel.NoCompression);
                        using (StreamWriter writer = new StreamWriter(newDocXmlEntry.Open()))
                        {
                            writer.Write(xmlContent);
                        }
                    }
                }
            }

            Console.WriteLine($"Processed: {Path.GetFileName(docFile)}");
        }

        Console.WriteLine("Batch embedding completed.");
    }

    private static void CreateMinimalExcel(string path)
    {
        // Create a very small, valid XLSX file using the Open Packaging Conventions.
        // The file will contain the minimal parts required for Excel to open it.
        using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
        using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Create))
        {
            // [Content_Types].xml
            var ct = zip.CreateEntry("[Content_Types].xml", CompressionLevel.NoCompression);
            using (var writer = new StreamWriter(ct.Open()))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
  <Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
</Types>");
            }

            // _rels/.rels
            var rels = zip.CreateEntry("_rels/.rels", CompressionLevel.NoCompression);
            using (var writer = new StreamWriter(rels.Open()))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
</Relationships>");
            }

            // xl/workbook.xml
            var workbook = zip.CreateEntry("xl/workbook.xml", CompressionLevel.NoCompression);
            using (var writer = new StreamWriter(workbook.Open()))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""
          xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
  <sheets>
    <sheet name=""Sheet1"" sheetId=""1"" r:id=""rId1""/>
  </sheets>
</workbook>");
            }

            // xl/_rels/workbook.xml.rels
            var wbRels = zip.CreateEntry("xl/_rels/workbook.xml.rels", CompressionLevel.NoCompression);
            using (var writer = new StreamWriter(wbRels.Open()))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>
</Relationships>");
            }

            // xl/worksheets/sheet1.xml
            var sheet = zip.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.NoCompression);
            using (var writer = new StreamWriter(sheet.Open()))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <sheetData/>
</worksheet>");
            }
        }
    }

    private static void CreateMinimalWord(string path, string text)
    {
        // Create a very small, valid DOCX file using the Open Packaging Conventions.
        using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
        using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Create))
        {
            // [Content_Types].xml
            var ct = zip.CreateEntry("[Content_Types].xml", CompressionLevel.NoCompression);
            using (var writer = new StreamWriter(ct.Open()))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/word/document.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml""/>
</Types>");
            }

            // _rels/.rels
            var rels = zip.CreateEntry("_rels/.rels", CompressionLevel.NoCompression);
            using (var writer = new StreamWriter(rels.Open()))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""word/document.xml""/>
</Relationships>");
            }

            // word/document.xml
            var doc = zip.CreateEntry("word/document.xml", CompressionLevel.NoCompression);
            using (var writer = new StreamWriter(doc.Open()))
            {
                writer.Write($@"<?xml version=""1.0"" encoding=""UTF-8""?>
<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:r>
        <w:t>{System.Security.SecurityElement.Escape(text)}</w:t>
      </w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>");
            }
        }
    }
}
