using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Configure find/replace to use a custom callback that inserts a shape.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ShapeReplacingHandler()
        };

        // Replace every occurrence of the placeholder [SHAPE] with a rectangle shape.
        doc.Range.Replace(new Regex(@"\[SHAPE\]"), string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Callback that inserts a shape at the location of each match.
    private class ShapeReplacingHandler : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match is inside a Run node; get its parent paragraph (optional, only needed if you want to move the builder to the paragraph).
            Paragraph paragraph = (Paragraph)args.MatchNode.ParentNode;

            // Create a DocumentBuilder for the document. Cast DocumentBase to Document because the constructor expects a Document instance.
            DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);

            // Move the builder to the position of the matched node (the placeholder text).
            builder.MoveTo(args.MatchNode);

            // Insert a rectangle shape (adjust type and size as required).
            builder.InsertShape(ShapeType.Rectangle, 100, 50);

            // Remove the original matched text node.
            args.MatchNode.Remove();

            // Skip the default text replacement because we have already handled the match.
            return ReplaceAction.Skip;
        }
    }
}
