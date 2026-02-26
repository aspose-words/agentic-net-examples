using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading for the OLE object section.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("OLE Object Example");

        // Return to normal style for the following content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Insert a description paragraph.
        builder.Writeln("The following shape embeds a ZIP archive as an OLE package.");

        // Load the binary data of the file to embed (e.g., a ZIP archive).
        // Adjust the path to point to an existing file on your system.
        string zipFilePath = @"MyDir\sample.zip";
        byte[] zipBytes = File.ReadAllBytes(zipFilePath);

        // Insert the OLE object from a memory stream.
        using (MemoryStream zipStream = new MemoryStream(zipBytes))
        {
            // progId "Package" indicates a generic OLE package.
            // asIcon = false inserts the object as its content (not as an icon).
            // presentation = null lets Aspose.Words choose a default presentation.
            Shape oleShape = builder.InsertOleObject(zipStream, "Package", false, null);

            // Optionally set the display name and file name for the embedded package.
            oleShape.OleFormat.OlePackage.FileName = "sample.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample ZIP Archive";
        }

        // Insert a line break before the next example.
        builder.InsertBreak(BreakType.LineBreak);

        // Insert the same OLE object but displayed as an icon with a custom caption.
        builder.Writeln("The same ZIP archive displayed as an icon:");
        using (MemoryStream zipStream = new MemoryStream(zipBytes))
        {
            // Insert as an icon (asIcon = true) and provide a custom caption.
            Shape oleIconShape = builder.InsertOleObject(zipStream, "Package", true, null);
            oleIconShape.OleFormat.OlePackage.FileName = "sample.zip";
            oleIconShape.OleFormat.OlePackage.DisplayName = "Sample ZIP Archive (Icon)";
            // The icon caption defaults to the file name; you can change it via the OLE package if needed.
        }

        // Save the document to a DOCX file.
        // Adjust the output path as required.
        doc.Save(@"ArtifactsDir\InsertOleObjectExample.docx");
    }
}
