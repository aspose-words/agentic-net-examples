using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ReplaceOleObjectExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Step 1: Insert a placeholder OLE object (a generic Package) so that
        // we have something to replace later.
        // -----------------------------------------------------------------
        using (MemoryStream dummyData = new MemoryStream())
        {
            // Write some text into the stream.
            using (StreamWriter writer = new StreamWriter(dummyData))
            {
                writer.Write("Placeholder OLE package data");
                writer.Flush();
                dummyData.Position = 0;

                // Insert the dummy OLE object. The returned Shape represents the OLE object.
                Shape oldOleShape = builder.InsertOleObject(dummyData, "Package", false, null);

                // -----------------------------------------------------------------
                // Step 2: Replace the placeholder OLE object with a new image OLE object.
                // -----------------------------------------------------------------
                // Move the builder cursor to the old OLE shape.
                builder.MoveTo(oldOleShape);

                // A 1x1 pixel PNG image (transparent) encoded in Base64.
                const string pngBase64 =
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XkWcAAAAASUVORK5CYII=";
                byte[] pngBytes = Convert.FromBase64String(pngBase64);

                // Insert the new image as an OLE object (embedded, not an icon).
                using (MemoryStream imageStream = new MemoryStream(pngBytes))
                {
                    builder.InsertOleObject(imageStream, "Package", false, null);
                }

                // Remove the original placeholder OLE shape now that the new one is inserted.
                oldOleShape.Remove();
            }
        }

        // Save the resulting document.
        string outputPath = "OutputDocument.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}
