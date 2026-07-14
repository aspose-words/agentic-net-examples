using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("Document with an OLE object inside a content control:");

        // Create a block‑level rich‑text content control.
        StructuredDocumentTag oleContainer = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "OleContainer",
            Tag = "OleCC"
        };
        // Append the content control to the body.
        doc.FirstSection.Body.AppendChild(oleContainer);

        // Add a paragraph inside the content control – the builder will insert content here.
        Paragraph innerParagraph = new Paragraph(doc);
        oleContainer.AppendChild(innerParagraph);
        builder.MoveTo(innerParagraph);

        // Prepare some sample data to embed as an OLE package.
        byte[] sampleData = System.Text.Encoding.UTF8.GetBytes("This is sample OLE package content.");
        using (MemoryStream dataStream = new MemoryStream(sampleData))
        {
            // Insert the OLE object (as a package) directly into the content control.
            Shape oleShape = builder.InsertOleObject(dataStream, "Package", false, null);

            // Optionally set display properties for the embedded package.
            oleShape.OleFormat.OlePackage.FileName = "sample.txt";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample Text File";
        }

        // Save the document as DOCX.
        string docxPath = Path.Combine(outputDir, "OleInContentControl.docx");
        doc.Save(docxPath);

        // Convert the document to PDF, ensuring OLE control images are rendered.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            UpdateOleControlImages = true
        };
        string pdfPath = Path.Combine(outputDir, "OleInContentControl.pdf");
        doc.Save(pdfPath, pdfOptions);
    }
}
