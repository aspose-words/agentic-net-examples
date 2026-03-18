using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Ole;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an OLE object (example: a ZIP package) as an icon.
        byte[] zipBytes = File.ReadAllBytes("cat001.zip");
        using (MemoryStream stream = new MemoryStream(zipBytes))
        {
            // InsertOleObject is the provided rule for OLE insertion.
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Optional: set package display properties.
            oleShape.OleFormat.OlePackage.FileName = "cat001.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "cat001.zip";

            // Retrieve the current display size of the OLE object (in points).
            double originalWidth = oleShape.Width;
            double originalHeight = oleShape.Height;

            // Adjust the size – for example, increase both dimensions by 150%.
            oleShape.Width = originalWidth * 1.5;
            oleShape.Height = originalHeight * 1.5;
        }

        // Save the document using the provided lifecycle rule.
        doc.Save("OleObjectResized.docx");
    }
}
