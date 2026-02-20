using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Drawing;

class ShapeReplacingHandler : IReplacingCallback
{
    // This method is called for each match found during the replace operation.
    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // Create a new picture shape.
        Shape shape = new Shape(args.MatchNode.Document, ShapeType.Image);
        // Set the image for the shape (replace with your own image path).
        shape.ImageData.SetImage("ReplacementImage.png");
        // Set desired size.
        shape.Width = 100;
        shape.Height = 100;
        // Make the shape inline so it behaves like a character.
        shape.WrapType = WrapType.Inline;

        // Insert the shape after the paragraph that contains the match.
        Paragraph paragraph = (Paragraph)args.MatchNode.ParentNode;
        paragraph.ParentNode.InsertAfter(shape, paragraph);

        // Remove the original matched text node.
        args.MatchNode.Remove();

        // Skip the default replace action because we have already performed the replacement.
        return ReplaceAction.Skip;
    }
}

class FindReplaceWithShapeExample
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Set up find/replace options with our custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new ShapeReplacingHandler();

        // Perform the replace: every occurrence of the placeholder "[SHAPE]" will be replaced by the shape.
        doc.Range.Replace(new Regex(@"\[SHAPE\]"), "", options);

        // Save the modified document.
        doc.Save("OutputDocument.docx");
    }
}
