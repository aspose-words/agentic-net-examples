using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Load the file that will be embedded as an OLE object.
        // In this example we embed a ZIP archive; replace the path with your own file.
        byte[] fileBytes = File.ReadAllBytes("cat001.zip");
        using (MemoryStream stream = new MemoryStream(fileBytes))
        {
            // Insert the OLE object from the stream.
            // Parameters:
            //   stream   – the data stream of the file to embed.
            //   "Package" – ProgID indicating a generic OLE package.
            //   true     – display the object as an icon.
            //   null     – no custom presentation image; use the default icon.
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Optionally set the file name and display name that appear when the object is opened.
            oleShape.OleFormat.OlePackage.FileName = "Package file name.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "Package display name.zip";
        }

        // Save the resulting document to disk.
        doc.Save("InsertOleObject.docx");
    }
}
