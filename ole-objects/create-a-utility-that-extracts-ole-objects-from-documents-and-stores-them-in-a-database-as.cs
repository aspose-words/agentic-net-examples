using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OleExtractor
{
    public static void Main()
    {
        // Prepare a sample file that will be embedded as an OLE object.
        const string sampleTextPath = "sample.txt";
        File.WriteAllText(sampleTextPath, "Hello from embedded OLE object!");

        // Create a new Word document and embed the sample file as an OLE object.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the OLE object (embedded, not linked, displayed as content).
        builder.InsertOleObject(sampleTextPath, false, false, null);
        // Save the document to disk – this uses the provided save rule.
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Load the document back – this uses the provided load rule.
        Document loadedDoc = new Document(docPath);

        // Directory to store extracted OLE objects (simulating a BLOB database).
        const string storageDir = "OleBlobs";
        Directory.CreateDirectory(storageDir);

        // Iterate through all shapes in the document.
        int oleIndex = 0;
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue; // Not an OLE object.

            // Retrieve raw OLE data.
            byte[] oleData = oleFormat.GetRawData();

            // Determine a file name for the extracted object.
            string suggestedName = oleFormat.SuggestedFileName;
            if (string.IsNullOrEmpty(suggestedName))
            {
                // Fallback to a generated name using suggested extension.
                string extension = oleFormat.SuggestedExtension ?? ".bin";
                suggestedName = $"OleObject_{oleIndex}{extension}";
            }

            // Save the OLE data to a file in the storage directory.
            string outputPath = Path.Combine(storageDir, suggestedName);
            File.WriteAllBytes(outputPath, oleData);
            Console.WriteLine($"Extracted OLE object saved to: {outputPath}");

            oleIndex++;
        }

        // Clean up temporary files used for the demonstration.
        File.Delete(sampleTextPath);
        File.Delete(docPath);
    }
}
