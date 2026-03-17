using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace CommentEventDemo
{
    // Implements INodeChangingCallback to receive notifications when nodes are inserted or removed.
    class CommentLogger : INodeChangingCallback
    {
        private readonly StringBuilder _log = new StringBuilder();

        // Called after a node has been inserted into the document.
        void INodeChangingCallback.NodeInserted(NodeChangingArgs args)
        {
            if (args.Node.NodeType == NodeType.Comment)
            {
                Comment comment = (Comment)args.Node;
                _log.AppendLine($"[Added] Author: {comment.Author}, Id: {comment.Id}, Date: {comment.DateTime}");
            }
        }

        // Called before a node is inserted. Not used here.
        void INodeChangingCallback.NodeInserting(NodeChangingArgs args) { }

        // Called after a node has been removed from the document.
        void INodeChangingCallback.NodeRemoved(NodeChangingArgs args)
        {
            if (args.Node.NodeType == NodeType.Comment)
            {
                Comment comment = (Comment)args.Node;
                _log.AppendLine($"[Removed] Author: {comment.Author}, Id: {comment.Id}, Date: {comment.DateTime}");
            }
        }

        // Called before a node is removed. Not used here.
        void INodeChangingCallback.NodeRemoving(NodeChangingArgs args) { }

        // Exposes the accumulated log.
        public string GetLog() => _log.ToString();
    }

    class Program
    {
        static void Main()
        {
            // Create a new document.
            Document doc = new Document();

            // Attach the custom logger to receive node change events.
            CommentLogger logger = new CommentLogger();
            doc.NodeChangingCallback = logger;

            // Build some content.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample paragraph.");

            // Add a comment to the current paragraph.
            Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
            comment.SetText("First comment.");
            builder.CurrentParagraph.AppendChild(comment);

            // Add a second comment.
            Comment secondComment = new Comment(doc, "Bob", "B", DateTime.Now);
            secondComment.SetText("Second comment.");
            builder.CurrentParagraph.AppendChild(secondComment);

            // Remove the first comment.
            comment.Remove();

            // Output the log of comment additions and removals.
            Console.WriteLine("Comment event log:");
            Console.WriteLine(logger.GetLog());

            // Save the document.
            doc.Save("CommentEvents.docx");
        }
    }
}
