using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    // Custom callback that logs comment insertions and removals.
    private class CommentLogger : INodeChangingCallback
    {
        void INodeChangingCallback.NodeInserted(NodeChangingArgs args)
        {
            if (args.Node.NodeType == NodeType.Comment)
            {
                var comment = (Comment)args.Node;
                Console.WriteLine($"[Log] Comment added – Author: {comment.Author}, Text: \"{comment.GetText().Trim()}\"");
            }
        }

        void INodeChangingCallback.NodeInserting(NodeChangingArgs args) { }

        void INodeChangingCallback.NodeRemoved(NodeChangingArgs args)
        {
            if (args.Node.NodeType == NodeType.Comment)
            {
                var comment = (Comment)args.Node;
                Console.WriteLine($"[Log] Comment removed – Author: {comment.Author}, Text: \"{comment.GetText().Trim()}\"");
            }
        }

        void INodeChangingCallback.NodeRemoving(NodeChangingArgs args) { }
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the first paragraph of the document.");

        // Attach the custom node‑changing callback.
        doc.NodeChangingCallback = new CommentLogger();

        // Add first comment.
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("Please review this paragraph.");
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.AppendChild(comment1);

        // Add second comment.
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Consider rephrasing the sentence.");
        firstParagraph.AppendChild(comment2);

        // Remove the first comment to trigger the removal log.
        comment1.Remove();

        // Save the document to verify that comments are persisted.
        doc.Save("CommentEvents.docx");
    }
}
