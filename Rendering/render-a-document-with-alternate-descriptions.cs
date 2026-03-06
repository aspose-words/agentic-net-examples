using System;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Render comments as annotations (alternate description) in the output PDF.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

        // Use fallback shapes for DrawingML objects to provide alternate visual representations.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            DmlRenderingMode = DmlRenderingMode.Fallback
        };

        // Save the document to PDF with the specified rendering options.
        doc.Save("Output.pdf", saveOptions);
    }
}
