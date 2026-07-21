using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class SetOlePackageFileName
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some dummy data to embed as an OLE package.
        // In a real scenario this could be the bytes of a ZIP, PDF, etc.
        byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("This is the content of the embedded file.");

        // Insert the OLE object from the stream.
        // progId "Package" indicates a generic OLE package.
        // asIcon = true so the object appears as an icon.
        using (MemoryStream stream = new MemoryStream(dummyData))
        {
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Set the original file name and display name for the OLE package.
            // This preserves the metadata when the document is opened in Word.
            oleShape.OleFormat.OlePackage.FileName = "EmbeddedFile.txt";
            oleShape.OleFormat.OlePackage.DisplayName = "EmbeddedFile.txt";
        }

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OlePackageExample.docx");
        doc.Save(outputPath);
    }
}
