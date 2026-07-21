using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OlePackageExample
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some dummy data to embed. Here we use a simple text file content.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("This is a sample OLE package content.");

        // Insert the OLE object using the legacy "Package" progId.
        // The object will be displayed as an icon.
        using (MemoryStream dataStream = new MemoryStream(dummyData))
        {
            Shape oleShape = builder.InsertOleObject(dataStream, "Package", true, null);

            // Configure the OLE package properties.
            oleShape.OleFormat.OlePackage.FileName = "SamplePackage.txt";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample Package";
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OlePackageExample.docx");
        doc.Save(outputPath);
    }
}
