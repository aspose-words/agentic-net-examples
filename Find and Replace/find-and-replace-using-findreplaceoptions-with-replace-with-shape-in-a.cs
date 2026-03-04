using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Drawing;

namespace FindReplaceWithShapeExample
{
    // Custom callback that inserts a shape at each match location.
    class ShapeReplacingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The document that contains the match is a DocumentBase; cast it to Document.
            Document doc = (Document)args.MatchNode.Document;
            // Create a builder for that document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Position the builder at the start of the matched node.
            builder.MoveTo(args.MatchNode);
            // Insert the desired shape (e.g., a rectangle 100x50 points).
            builder.InsertShape(ShapeType.Rectangle, 100, 50);
            // Remove the placeholder text by replacing it with an empty string.
            args.Replacement = string.Empty;
            // Perform the replacement (which will delete the placeholder).
            return ReplaceAction.Replace;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the source DOCX document.
            Document doc = new Document("Input.docx");

            // Configure find/replace options to use the custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new ShapeReplacingCallback()
            };

            // Replace every occurrence of the placeholder with a shape.
            // The actual replacement text is set to empty inside the callback.
            doc.Range.Replace("_PLACEHOLDER_", "", options);

            // Save the modified document.
            doc.Save("Output.docx");
        }
    }
}
