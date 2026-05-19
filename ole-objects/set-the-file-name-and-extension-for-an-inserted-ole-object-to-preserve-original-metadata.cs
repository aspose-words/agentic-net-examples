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
        // In a real scenario this could be any file (e.g., a ZIP archive).
        byte[] packageData = System.Text.Encoding.UTF8.GetBytes("Hello, this is the content of the OLE package.");

        // Insert the OLE object from the memory stream.
        using (MemoryStream stream = new MemoryStream(packageData))
        {
            // Insert as an icon (asIcon = true) with the generic "Package" progId.
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Preserve original file name and extension metadata.
            oleShape.OleFormat.OlePackage.FileName = "SamplePackage.txt";
            oleShape.OleFormat.OlePackage.DisplayName = "SamplePackage.txt";
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OlePackageExample.docx");
        doc.Save(outputPath);
    }
}
