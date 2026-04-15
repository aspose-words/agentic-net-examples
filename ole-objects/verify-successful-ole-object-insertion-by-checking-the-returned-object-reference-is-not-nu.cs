using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some dummy data to embed as an OLE object.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Hello OLE");
        using (MemoryStream dataStream = new MemoryStream(dummyData))
        {
            // Insert the OLE object. The method returns the Shape that contains the OLE object.
            Shape oleShape = builder.InsertOleObject(dataStream, "Package", false, null);

            // Verify that the returned Shape reference is not null.
            if (oleShape == null)
                throw new InvalidOperationException("InsertOleObject returned a null Shape.");

            // Verify that the Shape actually contains an OLE object.
            if (oleShape.OleFormat == null)
                throw new InvalidOperationException("The inserted Shape does not have an OleFormat.");

            // Optional: output verification result.
            Console.WriteLine("OLE object inserted successfully. Shape and OleFormat are valid.");
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectTest.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
