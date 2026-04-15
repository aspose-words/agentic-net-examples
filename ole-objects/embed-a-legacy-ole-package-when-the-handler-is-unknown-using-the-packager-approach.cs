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

        // Prepare some data to embed. Here we use a simple text file content.
        byte[] fileBytes = System.Text.Encoding.UTF8.GetBytes("This is the content of the embedded package.");

        // Insert the data as an OLE package (handler unknown) using the Packager approach.
        // The progId "Package" tells Aspose.Words to treat the data as a generic OLE package.
        // The object is inserted as an icon (asIcon = true). No custom presentation image is provided.
        using (MemoryStream stream = new MemoryStream(fileBytes))
        {
            Shape oleShape = builder.InsertOleObject(stream, "Package", true, null);

            // Set the package's file name and display name that will appear in Word.
            oleShape.OleFormat.OlePackage.FileName = "EmbeddedFile.txt";
            oleShape.OleFormat.OlePackage.DisplayName = "Embedded File.txt";
        }

        // Save the document to the local file system.
        doc.Save("OlePackage.docx");
    }
}
