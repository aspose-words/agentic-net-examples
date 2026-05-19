using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OleProgIdDemo
{
    public static void Main()
    {
        // Prepare a temporary folder for the demo files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeOleDemo");
        Directory.CreateDirectory(tempFolder);

        // Create a simple text file that will be embedded as an OLE object.
        string sampleFilePath = Path.Combine(tempFolder, "sample.txt");
        File.WriteAllText(sampleFilePath, "This is a sample text file for OLE embedding.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the text file as an OLE object (embedded, not linked, displayed as content).
        // Use the "Package" ProgID which is suitable for generic file packages.
        Shape oleShape = builder.InsertOleObject(
            sampleFilePath,   // file to embed
            "Package",        // ProgID
            false,            // isLinked
            false,            // asIcon
            null);            // no custom presentation image

        // Retrieve the ProgID of the inserted OLE object.
        string progId = oleShape.OleFormat.ProgId;

        // Log the ProgID to the console.
        Console.WriteLine($"Inserted OLE object's ProgId: {progId}");

        // Save the document to the temporary folder.
        string outputPath = Path.Combine(tempFolder, "OleDemo.docx");
        doc.Save(outputPath);
    }
}
