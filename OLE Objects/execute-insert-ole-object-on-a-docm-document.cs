using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph describing the OLE object.
        builder.Writeln("Embedded ZIP file as OLE object:");

        // Load the file that will be embedded (replace with your actual file path).
        byte[] zipBytes = File.ReadAllBytes(@"MyDir\sample.zip");
        using (MemoryStream zipStream = new MemoryStream(zipBytes))
        {
            // Insert the OLE object from the stream.
            // progId "Package" indicates a generic OLE package.
            // asIcon = true displays the object as an icon.
            // presentation = null lets Aspose.Words choose a default icon.
            Shape oleShape = builder.InsertOleObject(zipStream, "Package", true, null);

            // Set the file name and display name that Word will show for the OLE package.
            oleShape.OleFormat.OlePackage.FileName = "sample.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample ZIP Archive";
        }

        // Save the document as a macro‑enabled DOCM file.
        doc.Save(@"ArtifactsDir\InsertOleObject.docm");
    }
}
