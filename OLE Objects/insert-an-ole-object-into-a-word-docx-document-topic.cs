using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Read the file that will be embedded (e.g., a ZIP archive) into a memory stream.
        byte[] fileBytes = File.ReadAllBytes("Data/sample.zip");
        using (MemoryStream stream = new MemoryStream(fileBytes))
        {
            // Insert the OLE object as an icon.
            // Parameters: stream with data, ProgID "Package", display as icon, no custom presentation image.
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Set the file name and display name that Word will show for the embedded package.
            oleShape.OleFormat.OlePackage.FileName = "sample.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample ZIP Archive";
        }

        // Save the document to a DOCX file.
        doc.Save("Output/DocumentWithOleObject.docx");
    }
}
