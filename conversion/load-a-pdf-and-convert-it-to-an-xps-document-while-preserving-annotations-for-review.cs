using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample Word document, add a comment (annotation) and
        //    save it as a PDF file with comments rendered as annotations.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF with a review comment.");

        // Create a comment node and attach it to the current paragraph.
        Comment comment = new Comment(sourceDoc, "Reviewer", "RV", DateTime.Now);
        comment.SetText("Please review this paragraph.");
        builder.CurrentParagraph.AppendChild(comment);

        // Render comments as annotations (required for PDF).
        sourceDoc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2. Load the generated PDF file.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Convert the PDF to XPS while preserving annotations.
        // -----------------------------------------------------------------
        const string xpsPath = "output.xps";
        XpsSaveOptions xpsOptions = new XpsSaveOptions(); // default options preserve annotations
        pdfDoc.Save(xpsPath, xpsOptions);

        // -----------------------------------------------------------------
        // 4. Validate that the XPS file was created and contains data.
        // -----------------------------------------------------------------
        if (!File.Exists(xpsPath) || new FileInfo(xpsPath).Length == 0)
        {
            throw new InvalidOperationException("XPS conversion failed: output file is missing or empty.");
        }

        Console.WriteLine("Conversion succeeded. XPS file created at: " + Path.GetFullPath(xpsPath));
    }
}
