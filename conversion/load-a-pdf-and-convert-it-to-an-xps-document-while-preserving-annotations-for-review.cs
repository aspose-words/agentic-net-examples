using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string pdfFile = "sample.pdf";
        const string xpsFile = "sample.xps";

        // ---------------------------------------------------------------
        // Step 1: Create a sample PDF document with a comment (annotation).
        // ---------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add some content.
        builder.Writeln("This is a sample document.");

        // Create a comment node manually (InsertComment is not available in this version).
        Comment comment = new Comment(sourceDoc, "Reviewer", "RV", DateTime.Now);
        comment.SetText("Please review this paragraph.");
        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Ensure comments are saved as PDF annotations.
        sourceDoc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;
        // Rebuild layout after changing layout options.
        sourceDoc.UpdatePageLayout();

        // Save the document as PDF.
        sourceDoc.Save(pdfFile, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfFile))
            throw new InvalidOperationException("The PDF file was not created.");

        // ---------------------------------------------------------------
        // Step 2: Load the PDF and convert it to XPS while preserving annotations.
        // ---------------------------------------------------------------
        Document pdfDoc = new Document(pdfFile);

        // Use XpsSaveOptions to control XPS output if needed.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Save the loaded PDF as XPS.
        pdfDoc.Save(xpsFile, xpsOptions);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsFile))
            throw new InvalidOperationException("The XPS file was not created.");

        // Indicate successful conversion.
        Console.WriteLine($"PDF file '{pdfFile}' was successfully converted to XPS file '{xpsFile}'.");
    }
}
