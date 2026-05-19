using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OlePackageExample
{
    public static void Main()
    {
        // Prepare a simple byte array to act as the embedded file data.
        byte[] fileData = System.Text.Encoding.UTF8.GetBytes("This is sample content for the OLE package.");

        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the byte array as an OLE Package object (displayed as an icon).
        using (MemoryStream stream = new MemoryStream(fileData))
        {
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Access the OlePackage and modify its properties.
            OlePackage olePackage = oleShape.OleFormat.OlePackage;
            olePackage.FileName = "SampleData.txt";
            olePackage.DisplayName = "Sample Data Package";
        }

        // Save the document containing the OLE package.
        string firstPath = Path.Combine(Environment.CurrentDirectory, "OlePackageInserted.docx");
        doc.Save(firstPath);

        // Load the saved document.
        Document loadedDoc = new Document(firstPath);
        Shape loadedShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);

        // Read the OlePackage properties.
        OlePackage loadedPackage = loadedShape.OleFormat.OlePackage;
        Console.WriteLine("Original FileName: " + loadedPackage.FileName);
        Console.WriteLine("Original DisplayName: " + loadedPackage.DisplayName);

        // Modify the properties again.
        loadedPackage.FileName = "RenamedData.txt";
        loadedPackage.DisplayName = "Renamed Data Package";

        // Save the modified document.
        string secondPath = Path.Combine(Environment.CurrentDirectory, "OlePackageModified.docx");
        loadedDoc.Save(secondPath);
    }
}
