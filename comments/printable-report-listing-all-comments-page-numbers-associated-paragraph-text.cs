using System;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Tables;

class CommentReportGenerator
{
    static void Main()
    {
        // Load the source document that contains comments.
        Document srcDoc = new Document("InputWithComments.docx");

        // Ensure the layout model is up‑to‑date so page numbers are accurate.
        srcDoc.UpdatePageLayout();

        // LayoutCollector maps any node in the document to the page on which it starts.
        LayoutCollector layoutCollector = new LayoutCollector(srcDoc);

        // Create a new document that will hold the printable report.
        Document reportDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(reportDoc);

        // Title for the report.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln("Comments Report");
        builder.Writeln();

        // Iterate over all top‑level comments in the source document.
        NodeCollection commentNodes = srcDoc.GetChildNodes(NodeType.Comment, true);
        foreach (Comment comment in commentNodes)
        {
            // Skip replies – they have an ancestor comment.
            if (comment.Ancestor != null) continue;

            // Determine the page number where the comment starts.
            int pageNumber = layoutCollector.GetStartPageIndex(comment);

            // Retrieve the paragraph that contains the commented range.
            // The comment is usually preceded by a CommentRangeStart node.
            Node? commentRangeStart = comment.PreviousSibling;
            while (commentRangeStart != null && commentRangeStart.NodeType != NodeType.CommentRangeStart)
                commentRangeStart = commentRangeStart.PreviousSibling;

            Paragraph? associatedParagraph = null;
            if (commentRangeStart != null)
                associatedParagraph = (Paragraph?)commentRangeStart.GetAncestor(NodeType.Paragraph);

            string paragraphText = associatedParagraph != null
                ? associatedParagraph.GetText().Trim()
                : "[No associated paragraph]";

            // Write the information to the report.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln($"Page {pageNumber}");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"Paragraph: {paragraphText}");
            builder.Writeln($"Comment by {comment.Author} on {comment.DateTime:G}:");
            builder.Writeln(comment.GetText().Trim());
            builder.Writeln(new string('-', 40));
        }

        // Save the report document.
        reportDoc.Save("CommentsReport.docx");

        // Optional: open the generated report with the default associated application.
        // Printing directly from Aspose.Words is not available in .NET Core/.NET 5+.
        // You can print the document via an external viewer or convert to PDF first.
        Process.Start(new ProcessStartInfo("CommentsReport.docx") { UseShellExecute = true });
    }
}
