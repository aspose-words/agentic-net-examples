using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text to the document.
        builder.Writeln("This is a sample paragraph.");

        // Add a comment to the last paragraph.
        Comment comment = new Comment(doc, "John Doe", "J.D.", DateTime.Now);
        // The Comment object contains a Paragraph collection. Add the comment text to the first paragraph.
        comment.Paragraphs[0].AppendChild(new Run(doc, "Please review this sentence."));
        builder.CurrentParagraph.AppendChild(comment);

        // Iterate through all comments in the document.
        foreach (Comment c in doc.GetChildNodes(NodeType.Comment, true))
        {
            // Output comment details.
            Console.WriteLine($"Comment Id: {c.Id}");
            Console.WriteLine($"Author: {c.Author}");
            Console.WriteLine($"Date: {c.DateTime}");
            Console.WriteLine($"Text: {c.GetText().Trim()}");
            Console.WriteLine($"Done: {c.Done}");
            Console.WriteLine();

            // Mark the comment as done.
            c.Done = true;
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
