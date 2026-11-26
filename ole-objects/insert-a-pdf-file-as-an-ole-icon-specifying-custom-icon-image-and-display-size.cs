using System;
using System.IO;
using System.IO.Compression;
using System.Text;

public class Program
{
    public static void Main()
    {
        const string outputPath = "Output.docx";

        // Minimal PDF content (just a valid header and EOF)
        byte[] pdfBytes = Encoding.ASCII.GetBytes("%PDF-1.4\n1 0 obj\n<<>>\nendobj\ntrailer\n<<>>\n%%EOF");

        // Tiny 1x1 transparent PNG (used as the custom icon)
        const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO6X6ZcAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(pngBase64);

        // XML parts required by a minimal .docx with an OLE object
        string contentTypesXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Default Extension=""png"" ContentType=""image/png""/>
  <Default Extension=""pdf"" ContentType=""application/pdf""/>
  <Override PartName=""/word/document.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml""/>
  <Override PartName=""/word/_rels/document.xml.rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
</Types>";

        string relsXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""word/document.xml""/>
</Relationships>";

        string documentXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
            xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
            xmlns:v=""urn:schemas-microsoft-com:vml""
            xmlns:o=""urn:schemas-microsoft-com:office:office"">
  <w:body>
    <w:p>
      <w:r>
        <v:shape id=""OleObjectShape"" style=""width:100pt;height:100pt"" o:ole="""">
          <v:imagedata r:id=""rId2"" o:title=""PDF Icon""/>
          <o:OLEObject Type=""Package"" ProgID=""AcroExch.Document.DC"" ShapeID=""_x0000_i1025""
                       DrawAspect=""Icon"" ObjectID=""_1025"" r:id=""rId1""/>
        </v:shape>
      </w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>";

        string documentRelsXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"" Target=""embeddings/embedded.pdf""/>
  <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"" Target=""media/icon.png""/>
</Relationships>";

        // Create the .docx (zip package) and add all parts
        using (FileStream fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Create, leaveOpen: false))
        {
            // [Content_Types].xml
            AddEntry(archive, "[Content_Types].xml", contentTypesXml);

            // _rels/.rels
            AddEntry(archive, "_rels/.rels", relsXml);

            // word/document.xml
            AddEntry(archive, "word/document.xml", documentXml);

            // word/_rels/document.xml.rels
            AddEntry(archive, "word/_rels/document.xml.rels", documentRelsXml);

            // word/embeddings/embedded.pdf
            AddEntry(archive, "word/embeddings/embedded.pdf", pdfBytes);

            // word/media/icon.png
            AddEntry(archive, "word/media/icon.png", pngBytes);
        }

        Console.WriteLine($"Document created: {outputPath}");
    }

    private static void AddEntry(ZipArchive archive, string entryName, string content)
    {
        ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.NoCompression);
        using (StreamWriter writer = new StreamWriter(entry.Open(), new UTF8Encoding(false)))
        {
            writer.Write(content);
        }
    }

    private static void AddEntry(ZipArchive archive, string entryName, byte[] data)
    {
        ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.NoCompression);
        using (BinaryWriter writer = new BinaryWriter(entry.Open()))
        {
            writer.Write(data);
        }
    }
}
