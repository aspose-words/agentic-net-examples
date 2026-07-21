using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OlePackageExample
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the document that will contain the OLE package.
        string docPath = Path.Combine(outputDir, "OlePackageDocument.docx");

        // 1. Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 2. Create some sample data to embed as an OLE package.
        byte[] sampleData = System.Text.Encoding.UTF8.GetBytes("Hello, this is the content of the OLE package.");
        using (MemoryStream dataStream = new MemoryStream(sampleData))
        {
            // 3. Insert the OLE package into the document as an icon.
            //    ProgId "Package" indicates a generic OLE package.
            Shape oleShape = builder.InsertOleObject(dataStream, "Package", true, null);

            // 4. Access the OlePackage object to modify its properties.
            OlePackage olePackage = oleShape.OleFormat.OlePackage;
            olePackage.FileName = "Greeting.txt";
            olePackage.DisplayName = "Greeting File";

            // Optional: write the modified values to the console.
            Console.WriteLine("Inserted OLE package with:");
            Console.WriteLine($"  FileName   = {olePackage.FileName}");
            Console.WriteLine($"  DisplayName= {olePackage.DisplayName}");
        }

        // 5. Save the document containing the OLE package.
        doc.Save(docPath);
        Console.WriteLine($"Document saved to: {docPath}");

        // -----------------------------------------------------------------
        // 6. Load the saved document and read back the OLE package properties.
        Document loadedDoc = new Document(docPath);
        // Find the first shape that has an OleFormat (the OLE package we inserted).
        Shape loadedOleShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
        OlePackage loadedOlePackage = loadedOleShape.OleFormat.OlePackage;

        // 7. Output the properties read from the loaded document.
        Console.WriteLine("Loaded OLE package properties:");
        Console.WriteLine($"  FileName   = {loadedOlePackage.FileName}");
        Console.WriteLine($"  DisplayName= {loadedOlePackage.DisplayName}");

        // 8. Optionally, extract the raw data of the OLE package to a file.
        string extractedFilePath = Path.Combine(outputDir, loadedOlePackage.FileName);
        using (FileStream fileStream = new FileStream(extractedFilePath, FileMode.Create, FileAccess.Write))
        {
            // OleFormat.Save writes the embedded object's data to the provided stream.
            loadedOleShape.OleFormat.Save(fileStream);
        }
        Console.WriteLine($"Extracted OLE package data saved to: {extractedFilePath}");
    }
}
