using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define paths for the input document, output document, and the replacement image.
        string dataDir = Path.GetFullPath("Data");
        Directory.CreateDirectory(dataDir);
        string inputDocPath = Path.Combine(dataDir, "Input.docx");
        string outputDocPath = Path.Combine(dataDir, "Output.docx");
        string imagePath = Path.Combine(dataDir, "Image.png");

        // -----------------------------------------------------------------
        // Create a sample input document with an OLE object if it does not exist.
        // This ensures the example runs without external files.
        // -----------------------------------------------------------------
        if (!File.Exists(inputDocPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder sampleBuilder = new DocumentBuilder(sampleDoc);
            sampleBuilder.Writeln("Original OLE object:");

            // Insert a dummy OLE package (a simple text file) as the original object.
            byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Dummy content");
            using (MemoryStream ms = new MemoryStream(dummyData))
            {
                Shape oleShape = sampleBuilder.InsertOleObject(ms, "Package", true, null);
                oleShape.OleFormat.OlePackage.FileName = "Dummy.txt";
                oleShape.OleFormat.OlePackage.DisplayName = "Dummy.txt";
            }

            sampleDoc.Save(inputDocPath);
        }

        // -----------------------------------------------------------------
        // Create a simple PNG image if it does not exist.
        // The PNG data represents a 1x1 pixel transparent image.
        // -----------------------------------------------------------------
        if (!File.Exists(imagePath))
        {
            // Base64-encoded PNG (1x1 pixel, transparent)
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imagePath, pngBytes);
        }

        // -----------------------------------------------------------------
        // Load the existing document that contains the OLE object.
        // -----------------------------------------------------------------
        Document doc = new Document(inputDocPath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Find the first OLE object shape in the document.
        // -----------------------------------------------------------------
        Shape existingOleShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (existingOleShape != null && existingOleShape.ShapeType == ShapeType.OleObject)
        {
            // Remove the old OLE object.
            existingOleShape.Remove();

            // Insert a new OLE object that embeds the image file.
            // Parameters: fileName, isLinked (false), asIcon (false), presentation (null).
            builder.InsertOleObject(imagePath, false, false, null);
        }

        // -----------------------------------------------------------------
        // Save the modified document.
        // -----------------------------------------------------------------
        doc.Save(outputDocPath);
    }
}
