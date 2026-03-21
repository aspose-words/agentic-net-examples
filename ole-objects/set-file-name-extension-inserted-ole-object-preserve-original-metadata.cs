using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create sample data to embed as an OLE package.
        string fileName = "sample.txt";
        byte[] fileBytes = Encoding.UTF8.GetBytes("This is a sample embedded file.");

        // Insert the OLE object from a memory stream.
        using (MemoryStream stream = new MemoryStream(fileBytes))
        {
            // "Package" progId creates a generic OLE package.
            // The third argument (true) inserts the object as an icon.
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Preserve the original file name and extension in the OLE package metadata.
            oleShape.OleFormat.OlePackage.FileName = fileName;
            oleShape.OleFormat.OlePackage.DisplayName = fileName;
        }

        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "OlePackage.docx");

        // Save the document to disk.
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
