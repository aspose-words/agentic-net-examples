using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ExtractOleObject
{
    public static void Main()
    {
        // Create a simple text file in memory to embed as an OLE package.
        byte[] fileBytes = System.Text.Encoding.UTF8.GetBytes("This is the content of the embedded file.");
        using (MemoryStream embedStream = new MemoryStream(fileBytes))
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the OLE object from the stream.
            // progId "Package" denotes a generic OLE package.
            // asIcon = true to display it as an icon (optional).
            // presentation = null to use the default icon.
            Shape oleShape = builder.InsertOleObject(embedStream, "Package", true, null);

            // Access the OleFormat of the inserted shape.
            OleFormat oleFormat = oleShape.OleFormat;

            // Determine a file name for the extracted OLE data.
            string suggestedExtension = oleFormat.SuggestedExtension ?? ".bin";
            string outputFile = Path.Combine(Directory.GetCurrentDirectory(),
                                             "ExtractedOleObject" + suggestedExtension);

            // Save the OLE object's binary data to the file via a stream.
            using (FileStream fileStream = new FileStream(outputFile, FileMode.Create))
            {
                oleFormat.Save(fileStream);
            }

            // Optional: indicate completion (no interactive input required).
            Console.WriteLine("OLE object extracted to: " + outputFile);
        }
    }
}
