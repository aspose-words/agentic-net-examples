using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a source document and embed an OLE package containing simple text.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        byte[] packageData = System.Text.Encoding.UTF8.GetBytes("Hello from OLE package");
        using (MemoryStream packageStream = new MemoryStream(packageData))
        {
            // Insert the OLE object as a package.
            Shape sourceOleShape = srcBuilder.InsertOleObject(packageStream, "Package", false, null);
            // Optionally set display name for the package.
            sourceOleShape.OleFormat.OlePackage.FileName = "Sample.txt";
            sourceOleShape.OleFormat.OlePackage.DisplayName = "Sample.txt";
        }

        // Extract the OLE object's raw data into a memory stream.
        Shape oleShape = (Shape)sourceDoc.GetChild(NodeType.Shape, 0, true);
        MemoryStream extractedData = new MemoryStream();
        oleShape.OleFormat.Save(extractedData);
        extractedData.Position = 0; // Reset stream position for reading.

        // Create a target document and insert the cloned OLE object using the extracted data.
        Document targetDoc = new Document();
        DocumentBuilder targetBuilder = new DocumentBuilder(targetDoc);
        targetBuilder.Writeln("Cloned OLE object:");
        targetBuilder.InsertOleObject(extractedData, "Package", false, null);

        // Save the resulting document.
        targetDoc.Save("ClonedOle.docx");
    }
}
