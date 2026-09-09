using System;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using Aspose.Words.Replacing;
using Aspose.Words.Fields;
using Aspose.Words.Layout;

public class Program
{
    // Simple logger that writes messages to the console and to a log file.
    private static void Log(string message)
    {
        Console.WriteLine(message);
        File.AppendAllText("comment-events.log", $"{DateTime.Now:O} - {message}{Environment.NewLine}");
    }

    public static void Main()
    {
        // Ensure the log file starts fresh.
        if (File.Exists("comment-events.log"))
            File.Delete("comment-events.log");

        // Create a new blank document.
        Document doc = new Document();

        // Attach a callback that logs comment insertions and removals.
        doc.NodeChangingCallback = new CommentLogger(Log);

        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will hold the comment.
        builder.Writeln("This paragraph will have a comment attached.");

        // Create a comment instance with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Review this paragraph.");

        // Append the comment to the current paragraph.
        Paragraph paragraph = builder.CurrentParagraph;
        paragraph.AppendChild(comment);

        // Save the document after adding the comment.
        doc.Save("CommentEvents.docx");

        // Enumerate all comments to demonstrate retrieval (should be one).
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment c in comments)
        {
            Log($"Enumerated comment. Author: {c.Author}, Text: \"{c.GetText().Trim()}\"");
        }

        // Remove the comment using its Remove method.
        comment.Remove();

        // Save the document after removal.
        doc.Save("CommentEventsRemoved.docx");

        // Verify that no comments remain.
        var remaining = doc.GetChildNodes(NodeType.Comment, true)
                           .OfType<Comment>()
                           .ToList();

        Log($"Remaining comments count: {remaining.Count}");
    }
}

// Callback that logs when comment nodes are inserted or removed.
public class CommentLogger : INodeChangingCallback
{
    private readonly Action<string> _logAction;

    public CommentLogger(Action<string> logAction)
    {
        _logAction = logAction;
    }

    void INodeChangingCallback.NodeInserting(NodeChangingArgs args) { }

    void INodeChangingCallback.NodeInserted(NodeChangingArgs args)
    {
        if (args.Node.NodeType == NodeType.Comment && args.Node is Comment comment)
        {
            _logAction($"Comment added (callback). Author: {comment.Author}, Text: \"{comment.GetText().Trim()}\"");
        }
    }

    void INodeChangingCallback.NodeRemoving(NodeChangingArgs args) { }

    void INodeChangingCallback.NodeRemoved(NodeChangingArgs args)
    {
        if (args.Node.NodeType == NodeType.Comment && args.Node is Comment comment)
        {
            _logAction($"Comment removed (callback). Author: {comment.Author}");
        }
    }
}
