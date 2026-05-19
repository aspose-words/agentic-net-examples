using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Set up directories.
        string baseDir = Directory.GetCurrentDirectory();
        string dataDir = Path.Combine(baseDir, "Data");
        Directory.CreateDirectory(dataDir);

        // Paths for the temporary files.
        string sampleTextFile = Path.Combine(dataDir, "Sample.txt");
        string inputDoc = Path.Combine(dataDir, "OleObject.docx");
        string outputOle = Path.Combine(dataDir, "ExtractedOle.bin");

        // Create a simple text file that will be embedded as an OLE object.
        File.WriteAllText(sampleTextFile, "This is a sample text file for OLE embedding.");

        // -----------------------------------------------------------------
        // Create a new Word document and embed the text file as an OLE object.
        // -----------------------------------------------------------------
        Document doc = new Document();                     // Create a blank document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the OLE object (as a package) without displaying it as an icon.
        builder.InsertOleObject(sampleTextFile, "Package", false, false, null);
        // Save the document that now contains the OLE object.
        doc.Save(inputDoc);                               // Save the document to disk.

        // ---------------------------------------------------------------
        // Load the document, locate the OLE object, and extract its data.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(inputDoc);       // Load the previously saved document.
        // Find the first shape that contains an OLE object.
        Shape oleShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
        if (oleShape != null && oleShape.OleFormat != null)
        {
            OleFormat oleFormat = oleShape.OleFormat;

            // Option 1: Save directly to a file using the suggested extension.
            // string outputPath = Path.Combine(dataDir, "ExtractedOle" + oleFormat.SuggestedExtension);
            // oleFormat.Save(outputPath);

            // Option 2: Save via a stream (demonstrates the Save(Stream) overload).
            using (FileStream fs = new FileStream(outputOle, FileMode.Create))
            {
                oleFormat.Save(fs);                       // Save the OLE object's raw data to a binary file.
            }
        }
    }
}
