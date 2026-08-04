using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some sample data to embed as an OLE package (a simple text file).
        byte[] sampleData = System.Text.Encoding.UTF8.GetBytes("Sample OLE package content.");
        using (MemoryStream dataStream = new MemoryStream(sampleData))
        {
            // Insert the OLE object into the document as an icon.
            // The progId "Package" indicates a generic OLE package.
            builder.InsertOleObject(dataStream, "Package", true, null);
        }

        // Save the document to a temporary location (optional, just to have a file on disk).
        string docPath = Path.Combine(Path.GetTempPath(), "OleDocument.docx");
        doc.Save(docPath);

        // Reload the document to simulate a typical load scenario.
        Document loadedDoc = new Document(docPath);

        // Iterate through all shapes in the document.
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            // Check if the shape contains an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Retrieve the raw binary data of the OLE object.
            byte[] rawData = oleFormat.GetRawData();

            // Determine a suitable file extension using the SuggestedExtension property.
            string extension = oleFormat.SuggestedExtension ?? ".bin";

            // Create a temporary file name for the extracted OLE data.
            string tempFilePath = Path.Combine(Path.GetTempPath(),
                $"ExtractedOle_{Guid.NewGuid()}{extension}");

            // Write the raw data to the temporary file.
            File.WriteAllBytes(tempFilePath, rawData);

            // The temporary file now contains the OLE object's binary data and can be
            // used for external analysis. No further action is required.
        }

        // Clean up the temporary document file.
        if (File.Exists(docPath))
            File.Delete(docPath);
    }
}
