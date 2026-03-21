using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Create a temporary file to embed as an OLE object.
        string tempOlePath = Path.GetTempFileName();
        File.WriteAllText(tempOlePath, "Sample OLE content");

        try
        {
            // Insert the OLE object into the document.
            Shape oleShape = builder.InsertOleObject(tempOlePath, isLinked: false, asIcon: false, presentation: null);

            // Lock the shape's aspect ratio so it preserves its original proportions.
            oleShape.AspectRatioLocked = true;

            // Optional: set explicit size for the shape (in points).
            oleShape.Width = 200;
            oleShape.Height = 150;

            // Save the resulting document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectLocked.docx");
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
        finally
        {
            // Clean up the temporary file.
            if (File.Exists(tempOlePath))
                File.Delete(tempOlePath);
        }
    }
}
