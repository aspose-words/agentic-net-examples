using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a description paragraph.
        builder.Writeln("Embedded OLE object inside a Rich Text content control:");

        // Create a block‑level RichText content control.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        // Append the content control to the document body.
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // The content control must contain at least one paragraph to host the OLE object.
        Paragraph sdtParagraph = new Paragraph(doc);
        richTextSdt.AppendChild(sdtParagraph);

        // Move the builder to the paragraph inside the content control.
        builder.MoveTo(sdtParagraph);

        // Prepare a simple text file as the OLE package data.
        byte[] sampleData = System.Text.Encoding.UTF8.GetBytes("Hello from embedded OLE object!");
        using (MemoryStream oleStream = new MemoryStream(sampleData))
        {
            // Insert the OLE object (as a package) into the document.
            // progId "Package" indicates a generic OLE package.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", false, null);

            // Set display properties for the embedded package.
            oleShape.OleFormat.OlePackage.FileName = "Sample.txt";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample.txt";
        }

        // Save the document as DOCX (optional, for verification).
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);
        string docxPath = Path.Combine(outputDir, "OleInContentControl.docx");
        doc.Save(docxPath, SaveFormat.Docx);

        // Convert the document to PDF, ensuring OLE control images are rendered.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            UpdateOleControlImages = true
        };
        string pdfPath = Path.Combine(outputDir, "OleInContentControl.pdf");
        doc.Save(pdfPath, pdfOptions);
    }
}
