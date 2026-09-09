using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare dummy data to embed as an OLE package (e.g., a simple text file content).
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Sample content for OLE package.");

        // Insert the dummy data as an OLE Package object.
        using (MemoryStream stream = new MemoryStream(dummyData))
        {
            // "Package" progId indicates a generic OLE package.
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Set OLE package properties for later inspection.
            oleShape.OleFormat.OlePackage.FileName = "SamplePackage.txt";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample Package Display Name";
        }

        // Save the document to a temporary file.
        string docPath = Path.Combine(Path.GetTempPath(), "OlePackageDemo.docx");
        doc.Save(docPath);

        // Load the saved document.
        Document loadedDoc = new Document(docPath);

        // Find the first shape that contains an OLE object.
        Shape shapeWithOle = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
        if (shapeWithOle?.OleFormat?.OlePackage != null)
        {
            OlePackage olePackage = shapeWithOle.OleFormat.OlePackage;

            // Output OLE package information.
            Console.WriteLine("OLE Package Information:");
            Console.WriteLine($"  FileName   : {olePackage.FileName}");
            Console.WriteLine($"  DisplayName: {olePackage.DisplayName}");

            // Retrieve raw OLE data (the embedded file bytes) and display its length.
            byte[] rawData = shapeWithOle.OleFormat.GetRawData();
            Console.WriteLine($"  Raw data length: {rawData.Length} bytes");
        }
        else
        {
            Console.WriteLine("No OLE package found in the document.");
        }

        // Clean up the temporary document file.
        if (File.Exists(docPath))
        {
            File.Delete(docPath);
        }
    }
}
