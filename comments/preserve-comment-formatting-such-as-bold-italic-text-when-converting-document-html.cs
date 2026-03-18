using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PreserveCommentFormatting
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a regular paragraph.
        builder.Writeln("Sample paragraph.");

        // Create a comment that will contain bold and italic text.
        Comment comment = new Comment(doc, "Author", "A", DateTime.Now);

        // Bold run inside the comment.
        Run boldRun = new Run(doc, "Bold text");
        boldRun.Font.Bold = true;
        comment.AppendChild(boldRun);

        // Space between runs.
        comment.AppendChild(new Run(doc, " "));

        // Italic run inside the comment.
        Run italicRun = new Run(doc, "Italic text");
        italicRun.Font.Italic = true;
        comment.AppendChild(italicRun);

        // Attach the comment to the first paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Set HTML save options to preserve formatting (including comment formatting).
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Export round‑trip information keeps comment formatting in the HTML.
            ExportRoundtripInformation = true,
            // Optional: make the generated HTML more readable.
            PrettyFormat = true
        };

        // Save the document as HTML using the configured options.
        doc.Save("CommentPreserved.html", saveOptions);
    }
}
