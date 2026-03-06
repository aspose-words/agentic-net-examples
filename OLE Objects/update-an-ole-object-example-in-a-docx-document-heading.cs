using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleObjectExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading for the OLE object section.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Embedding OLE Objects");

        // Reset paragraph style for normal text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // -----------------------------------------------------------------
        // Example 1: Embed a ZIP archive as an OLE Package (displayed as an icon).
        // -----------------------------------------------------------------
        // Load the binary data of the ZIP file.
        byte[] zipBytes = File.ReadAllBytes(@"C:\Data\sample.zip");
        using (MemoryStream zipStream = new MemoryStream(zipBytes))
        {
            // Insert the OLE object. The object is embedded (asIcon = true) and no custom icon is provided.
            Shape oleShape = builder.InsertOleObject(zipStream, "Package", true, null);

            // Access the OlePackage to set display properties.
            OlePackage package = oleShape.OleFormat.OlePackage;
            package.FileName = "sample.zip";
            package.DisplayName = "Sample ZIP Archive"; // Caption shown under the icon.
        }

        // Add a line break between objects.
        builder.InsertBreak(BreakType.LineBreak);

        // -----------------------------------------------------------------
        // Example 2: Embed an Excel spreadsheet from file, displayed as content.
        // -----------------------------------------------------------------
        // Insert the Excel file as an embedded OLE object (asIcon = false).
        Shape excelShape = builder.InsertOleObject(@"C:\Data\Report.xlsx", false, false, null);

        // Optionally, modify OLE properties.
        excelShape.OleFormat.AutoUpdate = false; // Do not auto‑update the linked data.
        // IconCaption is read‑only; use OlePackage.DisplayName to set the caption.
        excelShape.OleFormat.OlePackage.DisplayName = "Embedded Excel Report";

        // Add another line break.
        builder.InsertBreak(BreakType.LineBreak);

        // -----------------------------------------------------------------
        // Example 3: Insert a PowerPoint presentation as an icon with a custom ICO file.
        // -----------------------------------------------------------------
        // Insert the OLE object as an icon, providing a custom icon file and caption.
        Shape pptShape = builder.InsertOleObjectAsIcon(
            @"C:\Data\Presentation.pptx",   // File to embed
            false,                           // Not a linked object
            @"C:\Icons\pptIcon.ico",       // Custom icon file
            "Presentation Overview");       // Icon caption

        // Update fields (if any) before saving.
        doc.UpdateFields();

        // Save the document to disk.
        doc.Save(@"C:\Output\OleObjectsExample.docx");
    }
}
