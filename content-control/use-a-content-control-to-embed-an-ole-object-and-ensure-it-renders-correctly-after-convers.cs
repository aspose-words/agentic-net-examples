using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert an introductory paragraph.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with an OLE object embedded inside a content control:");

        // Create a block‑level RichText content control.
        StructuredDocumentTag oleContainer = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "OleContentControl",
            Tag = "OleCC"
        };

        // The content control must contain at least one paragraph to host the OLE object.
        Paragraph innerParagraph = new Paragraph(doc);
        oleContainer.AppendChild(innerParagraph);
        doc.FirstSection.Body.AppendChild(oleContainer);

        // Move the builder to the paragraph inside the content control.
        builder.MoveTo(innerParagraph);

        // Prepare a simple text file as a stream to embed as an OLE package.
        byte[] sampleData = System.Text.Encoding.UTF8.GetBytes("Sample embedded text file content.");
        using (MemoryStream oleStream = new MemoryStream(sampleData))
        {
            // Insert the OLE object (as a package) into the document.
            // progId "Package" indicates a generic OLE package.
            // asIcon = false embeds the object directly; Word will render it as an icon in PDF.
            builder.InsertOleObject(oleStream, "Package", false, null);
        }

        // Save the document as DOCX (optional, for inspection).
        string docxPath = "OleInContentControl.docx";
        doc.Save(docxPath);

        // Convert the document to PDF. The OLE object should be rendered as an icon.
        string pdfPath = "OleInContentControl.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
