using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a dummy ZIP package in memory.
        byte[] zipBytes;
        using (MemoryStream zipStream = new MemoryStream())
        {
            using (var archive = new ZipArchive(zipStream, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("dummy.txt");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("This is a dummy file inside the ZIP package.");
                }
            }
            zipBytes = zipStream.ToArray();
        }

        // Insert the OLE object (ZIP package) as an icon.
        using (MemoryStream stream = new MemoryStream(zipBytes))
        {
            // InsertOleObject inserts the OLE object; "Package" is the ProgID for generic packages.
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Set package display properties.
            oleShape.OleFormat.OlePackage.FileName = "cat001.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "cat001.zip";

            // Retrieve the current display size of the OLE object (in points).
            double originalWidth = oleShape.Width;
            double originalHeight = oleShape.Height;

            // Adjust the size – for example, increase both dimensions by 150%.
            oleShape.Width = originalWidth * 1.5;
            oleShape.Height = originalHeight * 1.5;
        }

        // Save the document.
        doc.Save("OleObjectResized.docx");
    }
}
