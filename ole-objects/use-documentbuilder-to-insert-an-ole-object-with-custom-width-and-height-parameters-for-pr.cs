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

        // Write a description before the OLE object.
        builder.Writeln("Below is an embedded OLE object with custom size:");

        // Prepare some data to embed – here we use a simple text file content.
        byte[] fileBytes = System.Text.Encoding.UTF8.GetBytes("This is sample content for the OLE package.");
        using (MemoryStream oleStream = new MemoryStream(fileBytes))
        {
            // Insert the OLE object. "Package" is the ProgID for a generic OLE package.
            // asIcon = false (display the content), presentation = null (default icon if needed).
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", false, null);

            // Set custom width and height (in points). 1 point = 1/72 inch.
            oleShape.Width = 300;   // approx 4.17 inches
            oleShape.Height = 150;  // approx 2.08 inches
        }

        // Save the document to the file system.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OleObject.docx");
        doc.Save(outputPath);
    }
}
