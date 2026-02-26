using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Replacing;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace to use a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new ShapeReplacingCallback();

        // Replace the placeholder "[SHAPE]" with a rectangle shape.
        // The replacement string is empty because the callback inserts the shape.
        doc.Range.Replace(new Regex(@"\[SHAPE\]"), string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}

// Callback that replaces each found placeholder with a shape.
class ShapeReplacingCallback : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Obtain the document containing the match. The Document property returns DocumentBase,
        // so we need to cast it to Document.
        Document doc = (Document)args.MatchNode.Document;

        // Position the builder at the match node.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(args.MatchNode);

        // Insert a rectangle shape (100x50 points).
        builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Remove the original placeholder text node.
        args.MatchNode.Remove();

        // Skip the default replacement since we handled it.
        return ReplaceAction.Skip;
    }
}
