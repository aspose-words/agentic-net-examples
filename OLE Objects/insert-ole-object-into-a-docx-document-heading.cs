using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a heading for the OLE object section.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Embedded OLE Object Example");

        // Return to normal paragraph style for subsequent content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Path to the file that will be embedded as an OLE object (e.g., a ZIP archive).
        string zipFilePath = @"C:\Data\sample.zip";

        // Open the file as a stream and insert it as an OLE object displayed as an icon.
        using (FileStream zipStream = new FileStream(zipFilePath, FileMode.Open, FileAccess.Read))
        {
            // progId "Package" indicates a generic OLE package.
            // The third argument (true) specifies that the object should be shown as an icon.
            // The fourth argument (null) lets Aspose.Words choose a default icon.
            Shape oleShape = builder.InsertOleObject(zipStream, "Package", true, null);

            // Set custom file name and display name for the embedded package.
            oleShape.OleFormat.OlePackage.FileName = Path.GetFileName(zipFilePath);
            oleShape.OleFormat.OlePackage.DisplayName = "Sample Zip Archive";
        }

        // Save the resulting document to disk.
        doc.Save(@"C:\Output\OleObjectDemo.docx");
    }
}
