using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample PDF with a comment annotation.
        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample PDF with annotation.");

        // Add a comment (annotation) to the paragraph.
        Comment comment = new Comment(doc, "Author", "AU", DateTime.Now);
        comment.SetText("Sample annotation");
        builder.CurrentParagraph.AppendChild(comment);

        // Ensure comments are saved as annotations in the PDF.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF back as a Word document.
        Document pdfDoc = new Document(pdfPath);

        // Convert the loaded PDF to XPS while preserving annotations.
        string xpsPath = Path.Combine(outputDir, "sample.xps");
        XpsSaveOptions xpsOptions = new XpsSaveOptions(); // SaveFormat defaults to Xps.
        pdfDoc.Save(xpsPath, xpsOptions);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsPath))
        {
            throw new InvalidOperationException("XPS conversion failed: output file not found.");
        }

        // Indicate successful conversion.
        Console.WriteLine($"PDF successfully converted to XPS: {xpsPath}");
    }
}
