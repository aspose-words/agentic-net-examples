using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OlePackageExample
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the document that will contain the OLE package.
        string docPath = Path.Combine(outputDir, "OlePackageDocument.docx");

        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create some dummy data to embed as an OLE package (e.g., a simple text file content).
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("This is the content of the legacy OLE package.");

        // Insert the OLE package into the document.
        using (MemoryStream dataStream = new MemoryStream(dummyData))
        {
            // "Package" progId indicates a generic OLE package.
            // Insert as an icon (asIcon = true) with no custom presentation image.
            Shape oleShape = builder.InsertOleObject(dataStream, "Package", true, null);

            // Access the OlePackage object to modify its properties.
            OlePackage olePackage = oleShape.OleFormat.OlePackage;
            if (olePackage != null)
            {
                // Set custom file name and display name for the package.
                olePackage.FileName = "SamplePackage.txt";
                olePackage.DisplayName = "Sample Package Display Name.txt";
            }
        }

        // Save the document containing the OLE package.
        doc.Save(docPath);

        // Load the document back to demonstrate reading the OLE package properties.
        Document loadedDoc = new Document(docPath);
        Shape loadedShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
        OlePackage loadedPackage = loadedShape.OleFormat.OlePackage;

        // Output the stored properties to the console.
        if (loadedPackage != null)
        {
            Console.WriteLine("OLE Package FileName: " + loadedPackage.FileName);
            Console.WriteLine("OLE Package DisplayName: " + loadedPackage.DisplayName);
        }
        else
        {
            Console.WriteLine("No OLE package found in the loaded document.");
        }
    }
}
