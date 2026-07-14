using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some dummy data to embed as an OLE package.
        byte[] packageData = Encoding.UTF8.GetBytes("Sample OLE package content");
        using (MemoryStream packageStream = new MemoryStream(packageData))
        {
            // Insert the OLE package into the document as an icon.
            Shape oleShape = builder.InsertOleObject(packageStream, "Package", true, null);

            // Modify the OLE package properties.
            oleShape.OleFormat.OlePackage.FileName = "SamplePackage.txt";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample Package Display";

            // Read back the modified properties and output them.
            Console.WriteLine("Modified OLE Package properties:");
            Console.WriteLine($"FileName   : {oleShape.OleFormat.OlePackage.FileName}");
            Console.WriteLine($"DisplayName: {oleShape.OleFormat.OlePackage.DisplayName}");
        }

        // Save the document to the file system.
        doc.Save("OlePackageExample.docx");
    }
}
