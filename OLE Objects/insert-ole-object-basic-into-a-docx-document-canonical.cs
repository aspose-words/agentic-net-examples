using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Folder that contains the file to embed.
        string dataDir = @"C:\Data\";          // <-- change to your folder
        // Path where the resulting DOCX will be saved.
        string outputPath = @"C:\Output\OleObject.docx"; // <-- change to your folder

        // Load the file that will be embedded (e.g., a ZIP archive) into a byte array.
        byte[] fileBytes = File.ReadAllBytes(Path.Combine(dataDir, "sample.zip"));

        // Create a memory stream from the file bytes.
        using (MemoryStream fileStream = new MemoryStream(fileBytes))
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder to add content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Optional: add a description before the OLE object.
            builder.Writeln("Embedded OLE Package:");

            // Insert the OLE object from the stream.
            // Parameters:
            //   stream   – the data stream of the file.
            //   progId   – "Package" indicates a generic OLE package.
            //   asIcon   – true to display the object as an icon.
            //   presentation – null to use the default icon.
            Shape oleShape = builder.InsertOleObject(fileStream, "Package", true, null);

            // Set the file name and display name that Word will show.
            oleShape.OleFormat.OlePackage.FileName = "sample.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample ZIP";

            // Save the document to the specified path.
            doc.Save(outputPath);
        }

        Console.WriteLine("Document saved to: " + outputPath);
    }
}
