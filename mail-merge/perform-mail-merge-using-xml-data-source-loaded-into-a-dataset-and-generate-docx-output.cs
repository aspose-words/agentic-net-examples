using System;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Text;

public class MailMergeExample
{
    public static void Main()
    {
        // Sample XML data source
        string xmlData = @"
<Customers>
  <Customer>
    <FirstName>John</FirstName>
    <LastName>Doe</LastName>
    <Email>john@example.com</Email>
  </Customer>
  <Customer>
    <FirstName>Jane</FirstName>
    <LastName>Smith</LastName>
    <Email>jane@example.com</Email>
  </Customer>
</Customers>";

        // Load XML into a DataSet
        DataSet ds = new DataSet();
        using (StringReader sr = new StringReader(xmlData))
        {
            ds.ReadXml(sr);
        }

        // Build the WordprocessingML document (document.xml)
        StringBuilder docBuilder = new StringBuilder();
        docBuilder.Append(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
        docBuilder.Append(@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">");
        docBuilder.Append(@"<w:body>");
        docBuilder.Append(@"<w:p><w:r><w:t>Customer List</w:t></w:r></w:p>");
        docBuilder.Append(@"<w:tbl>");

        // Header row
        docBuilder.Append(@"<w:tr>");
        docBuilder.Append(Cell("First Name"));
        docBuilder.Append(Cell("Last Name"));
        docBuilder.Append(Cell("Email"));
        docBuilder.Append(@"</w:tr>");

        // Data rows
        foreach (DataRow row in ds.Tables["Customer"].Rows)
        {
            docBuilder.Append(@"<w:tr>");
            docBuilder.Append(Cell(row["FirstName"].ToString()));
            docBuilder.Append(Cell(row["LastName"].ToString()));
            docBuilder.Append(Cell(row["Email"].ToString()));
            docBuilder.Append(@"</w:tr>");
        }

        docBuilder.Append(@"</w:tbl>");
        // End of body
        docBuilder.Append(@"<w:sectPr>");
        docBuilder.Append(@"<w:pgSz w:w=""12240"" w:h=""15840""/>");
        docBuilder.Append(@"<w:pgMar w:top=""1440"" w:right=""1440"" w:bottom=""1440"" w:left=""1440""/>");
        docBuilder.Append(@"</w:sectPr>");
        docBuilder.Append(@"</w:body>");
        docBuilder.Append(@"</w:document>");

        string documentXml = docBuilder.ToString();

        // [Content_Types].xml
        string contentTypesXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/word/document.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml""/>
</Types>";

        // _rels/.rels
        string relsXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""word/document.xml""/>
</Relationships>";

        // Create the DOCX file (ZIP package)
        string outputPath = "Customers.docx";
        using (FileStream fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Create))
        {
            // Add [Content_Types].xml
            var ctEntry = zip.CreateEntry("[Content_Types].xml", CompressionLevel.NoCompression);
            using (var writer = new StreamWriter(ctEntry.Open(), Encoding.UTF8))
            {
                writer.Write(contentTypesXml);
            }

            // Add _rels/.rels
            var relsEntry = zip.CreateEntry("_rels/.rels", CompressionLevel.NoCompression);
            using (var writer = new StreamWriter(relsEntry.Open(), Encoding.UTF8))
            {
                writer.Write(relsXml);
            }

            // Add word/document.xml
            var docEntry = zip.CreateEntry("word/document.xml", CompressionLevel.NoCompression);
            using (var writer = new StreamWriter(docEntry.Open(), Encoding.UTF8))
            {
                writer.Write(documentXml);
            }
        }

        // Indicate completion (no interactive output required)
        Console.WriteLine($"DOCX file generated: {Path.GetFullPath(outputPath)}");
    }

    private static string Cell(string text)
    {
        return $@"<w:tc><w:p><w:r><w:t>{EscapeXml(text)}</w:t></w:r></w:p></w:tc>";
    }

    private static string EscapeXml(string text)
    {
        return System.Security.SecurityElement.Escape(text);
    }
}
