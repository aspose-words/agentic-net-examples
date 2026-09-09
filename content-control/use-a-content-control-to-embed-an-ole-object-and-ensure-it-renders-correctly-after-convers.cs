using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("Document with an OLE object inside a content control:");
        builder.Writeln();

        // Create a block‑level rich‑text content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "OleContentControl",
            Tag = "OleCC"
        };

        // The content control must contain at least one paragraph.
        Paragraph sdtParagraph = new Paragraph(doc);
        sdt.AppendChild(sdtParagraph);

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(sdt);

        // Move the builder cursor inside the newly created paragraph.
        builder.MoveTo(sdtParagraph);

        // Prepare a simple text file in memory to embed as an OLE package.
        byte[] oleData = Encoding.UTF8.GetBytes("This is the embedded OLE package content.");
        using (MemoryStream oleStream = new MemoryStream(oleData))
        {
            // Insert the OLE object. ProgId "Package" denotes a generic OLE package.
            // asIcon = false means the object will be displayed as its content.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", false, null);

            // Optionally set a display name for the OLE package.
            oleShape.OleFormat.OlePackage.FileName = "Sample.txt";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample.txt";
        }

        // Save the document as PDF. The OLE object should be rendered correctly.
        doc.Save("OleInContentControl.pdf", SaveFormat.Pdf);
    }
}
