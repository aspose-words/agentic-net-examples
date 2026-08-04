using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a simple byte array to act as the content of the legacy OLE package.
        byte[] packageData = System.Text.Encoding.UTF8.GetBytes("This is the content of a legacy OLE package.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE package into the document as an icon.
        using (MemoryStream stream = new MemoryStream(packageData))
        {
            // "Package" progId indicates a generic OLE package.
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Access the OlePackage and set its properties.
            OlePackage olePackage = oleShape.OleFormat.OlePackage;
            olePackage.FileName = "SamplePackage.txt";
            olePackage.DisplayName = "Sample Package Display Name.txt";
        }

        // Save the document containing the OLE package.
        string originalPath = "OlePackageDemo.docx";
        doc.Save(originalPath);

        // Load the saved document.
        Document loadedDoc = new Document(originalPath);

        // Find the first shape that contains an OLE object.
        Shape shapeWithOle = null;
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.OleFormat != null && shape.OleFormat.OlePackage != null)
            {
                shapeWithOle = shape;
                break;
            }
        }

        if (shapeWithOle != null)
        {
            OlePackage loadedPackage = shapeWithOle.OleFormat.OlePackage;

            // Read and display the current properties.
            Console.WriteLine("Original FileName: " + loadedPackage.FileName);
            Console.WriteLine("Original DisplayName: " + loadedPackage.DisplayName);

            // Modify the properties.
            loadedPackage.FileName = "ModifiedPackage.txt";
            loadedPackage.DisplayName = "Modified Package Display Name.txt";

            // Save the modified document.
            string modifiedPath = "OlePackageDemoModified.docx";
            loadedDoc.Save(modifiedPath);

            // Output confirmation.
            Console.WriteLine("Modified OLE package properties saved to " + modifiedPath);
        }
        else
        {
            Console.WriteLine("No OLE package found in the document.");
        }
    }
}
