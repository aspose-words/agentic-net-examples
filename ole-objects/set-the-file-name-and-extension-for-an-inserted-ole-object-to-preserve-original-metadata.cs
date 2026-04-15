using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some dummy data to embed as an OLE package.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("This is a dummy OLE package content.");

        // Insert the OLE object from the memory stream.
        using (MemoryStream stream = new MemoryStream(dummyData))
        {
            // Insert as an icon; progId "Package" indicates a generic OLE package.
            Shape shape = builder.InsertOleObject(stream, "Package", true, null);

            // Preserve original file name and extension metadata.
            shape.OleFormat.OlePackage.FileName = "DummyPackage.zip";
            shape.OleFormat.OlePackage.DisplayName = "DummyPackage.zip";
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OlePackage.docx");
        doc.Save(outputPath);
    }
}
