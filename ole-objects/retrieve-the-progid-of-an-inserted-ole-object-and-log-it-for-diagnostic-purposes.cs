using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary file that will be used as the source of the OLE object.
        string tempFilePath = Path.Combine(Path.GetTempPath(), "SampleOleSource.txt");
        File.WriteAllText(tempFilePath, "This is sample OLE content.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE object.
        // The overload used here is:
        // InsertOleObject(string fileName, string progId, bool isLink, bool isIcon, Stream iconStream)
        // We insert a non‑linked object, do not display an icon, and pass null for the icon stream.
        builder.InsertOleObject(tempFilePath, "Package", false, false, null);

        // Retrieve and display the ProgId of the inserted OLE object.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.ShapeType == ShapeType.OleObject && shape.OleFormat != null)
            {
                string progId = shape.OleFormat.ProgId;
                Console.WriteLine($"Inserted OLE object's ProgId: {progId}");
            }
        }

        // Save the document (optional, demonstrates lifecycle usage).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Clean up the temporary file.
        if (File.Exists(tempFilePath))
        {
            File.Delete(tempFilePath);
        }
    }
}
