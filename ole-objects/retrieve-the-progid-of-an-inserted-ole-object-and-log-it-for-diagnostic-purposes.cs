using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OleProgIdExample
{
    public static void Main()
    {
        // Create a temporary text file to embed as an OLE object.
        string tempFilePath = Path.Combine(Path.GetTempPath(), "SampleText.txt");
        File.WriteAllText(tempFilePath, "This is a sample text file for OLE embedding.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the temporary file as an embedded OLE object (not as an icon).
        // Parameters: file name, isLinked = false, asIcon = false, presentation = null.
        builder.InsertOleObject(tempFilePath, false, false, null);

        // Retrieve the first shape in the document, which should be the OLE object we just inserted.
        Shape oleShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (oleShape != null && oleShape.OleFormat != null)
        {
            // Access the OleFormat of the shape and get its ProgId.
            OleFormat oleFormat = oleShape.OleFormat;
            string progId = oleFormat.ProgId;

            // Log the ProgId to the console.
            Console.WriteLine($"Inserted OLE object's ProgId: {progId}");
        }
        else
        {
            Console.WriteLine("No OLE object found in the document.");
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleProgIdExample.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");

        // Clean up the temporary file.
        if (File.Exists(tempFilePath))
        {
            File.Delete(tempFilePath);
        }
    }
}
