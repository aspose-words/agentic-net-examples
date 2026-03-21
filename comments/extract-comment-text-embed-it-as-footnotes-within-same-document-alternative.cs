using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

class CommentToFootnote
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text.
        builder.Writeln("This is a sample paragraph.");

        // Add a comment to the paragraph.
        Comment comment = new Comment(doc, "Author", "A", DateTime.Now);
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, "This is a comment that will become a footnote."));
        comment.AppendChild(commentParagraph);
        builder.CurrentParagraph.AppendChild(comment);

        // Process all comments and convert them to footnotes.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);
        foreach (Comment c in commentNodes)
        {
            string commentText = c.GetText().Trim();
            builder.MoveTo(c);
            builder.InsertFootnote(FootnoteType.Footnote, commentText);
        }

        // Save the modified document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OutputWithFootnotes.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
