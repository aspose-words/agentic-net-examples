using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OleInsertionDemo
{
    public static void Main()
    {
        // Prepare a temporary folder for the demo files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "OleInsertionDemo");
        Directory.CreateDirectory(tempFolder);

        // Create a simple text file that will be inserted as an OLE object.
        string sampleFilePath = Path.Combine(tempFolder, "Sample.txt");
        File.WriteAllText(sampleFilePath, "This is a sample text file for OLE insertion.");

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the text file as an embedded OLE object (not linked, not displayed as an icon).
        Shape oleShape = builder.InsertOleObject(sampleFilePath, false, false, null);

        // Verify that the insertion returned a non‑null Shape and that it contains an OleFormat.
        bool insertionSuccessful = oleShape != null && oleShape.OleFormat != null;

        // Output the verification result.
        Console.WriteLine("OLE insertion successful: " + insertionSuccessful);

        // Save the document to a temporary file.
        string outputPath = Path.Combine(tempFolder, "OleDocument.docx");
        doc.Save(outputPath);

        // Clean up the temporary sample file (the document remains for inspection if needed).
        // File.Delete(sampleFilePath);
    }
}
