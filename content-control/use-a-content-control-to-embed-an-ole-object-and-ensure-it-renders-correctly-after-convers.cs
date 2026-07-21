using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an introductory paragraph.
        builder.Writeln("Document with an OLE object embedded inside a content control:");

        // Create a block‑level rich‑text content control.
        StructuredDocumentTag oleContainer = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "OleContainer",
            Tag = "OleCC"
        };

        // Add a paragraph that will hold the OLE object.
        Paragraph innerParagraph = new Paragraph(doc);
        oleContainer.AppendChild(innerParagraph);

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(oleContainer);

        // Position the builder inside the paragraph of the content control.
        builder.MoveTo(innerParagraph);

        // Prepare simple data to embed as an OLE package.
        byte[] oleData = System.Text.Encoding.UTF8.GetBytes("Hello OLE content");
        using (MemoryStream oleStream = new MemoryStream(oleData))
        {
            // Insert the OLE object (as a package) directly into the paragraph.
            // asIcon = false so the object is displayed as its content.
            builder.InsertOleObject(oleStream, "Package", false, null);
        }

        // Save the document as DOCX for reference.
        doc.Save("OleInContentControl.docx");

        // Save the document as PDF, ensuring OLE control images are rendered.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            UpdateOleControlImages = true
        };
        doc.Save("OleInContentControl.pdf", pdfOptions);
    }
}
