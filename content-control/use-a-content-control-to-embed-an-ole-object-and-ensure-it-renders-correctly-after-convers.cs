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

        // Add introductory text.
        builder.Writeln("Document with OLE object inside a content control:");

        // Insert a block‑level rich‑text content control.
        StructuredDocumentTag sdt = builder.InsertStructuredDocumentTag(SdtType.RichText);
        sdt.Title = "OleContentControl";
        sdt.Tag = "OleCC";

        // The builder is now positioned inside the content control.
        // Prepare a simple text file in memory to embed as an OLE package.
        byte[] oleData = System.Text.Encoding.UTF8.GetBytes("Hello from embedded OLE object!");
        using (MemoryStream oleStream = new MemoryStream(oleData))
        {
            // Insert the OLE object (as a shape) inside the content control.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", asIcon: false, presentation: null);

            // Set file name and display name for the embedded package.
            if (oleShape?.OleFormat?.OlePackage != null)
            {
                oleShape.OleFormat.OlePackage.FileName = "Sample.txt";
                oleShape.OleFormat.OlePackage.DisplayName = "Sample.txt";
            }
        }

        // Add concluding text.
        builder.Writeln();
        builder.Writeln("End of document.");

        // Save the document as DOCX.
        const string docPath = "OleInContentControl.docx";
        doc.Save(docPath);

        // Load the saved DOCX and convert it to PDF.
        Document pdfDoc = new Document(docPath);

        // Ensure OLE control images are updated during PDF conversion.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            UpdateOleControlImages = true
        };

        const string pdfPath = "OleInContentControl.pdf";
        pdfDoc.Save(pdfPath, pdfOptions);
    }
}
