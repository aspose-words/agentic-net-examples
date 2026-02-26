using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleInHtml
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert HTML content into the document.
        string html = "<h2>Embedded OLE Object Example</h2><p>This document contains an OLE package.</p>";
        builder.InsertHtml(html);

        // Add a paragraph break before inserting the OLE object.
        builder.InsertParagraph();

        // Load a file (e.g., a ZIP archive) into a memory stream.
        byte[] fileBytes = File.ReadAllBytes("sample.zip"); // Adjust the path to your file.
        using (MemoryStream oleStream = new MemoryStream(fileBytes))
        {
            // Insert the OLE object as an icon. "Package" is the ProgID for generic OLE packages.
            // asIcon = true displays the object as an icon; presentation = null uses the default icon.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", true, null);

            // If you need a custom caption, set the shape's AlternativeText or Title instead of IconCaption
            // because IconCaption is read‑only.
            oleShape.AlternativeText = "Sample ZIP Archive";
        }

        // Save the resulting document.
        doc.Save("OleInHtml.docx");
    }
}
