using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for the sample files.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // Paths for the intermediate DOCX and the final PDF.
        string docPath = Path.Combine(outputFolder, "SampleWithComments.docx");
        string pdfPath = Path.Combine(outputFolder, "SampleWithComments.pdf");

        // -------------------------------------------------
        // 1. Create a DOCX document and add a comment.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text that will have a comment attached.
        builder.Writeln("This paragraph contains a comment that should appear in the PDF.");

        // Create a comment, set its metadata and text, and attach it to the current paragraph.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
        comment.SetText("Please review this paragraph for accuracy.");
        builder.CurrentParagraph.AppendChild(comment);

        // Save the DOCX file (this simulates an existing Word document).
        doc.Save(docPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 2. Load the DOCX file, configure comment rendering, and convert to PDF.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // For PDF output we want comments to be shown as annotations.
        loadedDoc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

        // After changing layout options the page layout must be rebuilt.
        loadedDoc.UpdatePageLayout();

        // Save the document as PDF. All comments will be visible in the resulting file.
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);
    }
}
