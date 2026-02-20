using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace FindReplaceWithParagraphExample
{
    // Custom callback that replaces a found match with a new paragraph.
    class ReplaceWithParagraphHandler : IReplacingCallback
    {
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The node that contains the start of the match (usually a Run).
            // Its parent is the paragraph that we want to replace.
            Paragraph originalParagraph = (Paragraph)args.MatchNode.ParentNode;

            // Create a new paragraph in the same document.
            Paragraph newParagraph = new Paragraph(originalParagraph.Document);
            // Add the desired text to the new paragraph.
            newParagraph.AppendChild(new Run(originalParagraph.Document) { Text = "This is the replacement paragraph." });

            // Insert the new paragraph after the original one.
            originalParagraph.ParentNode.InsertAfter(newParagraph, originalParagraph);
            // Remove the original paragraph that contained the match.
            originalParagraph.Remove();

            // Skip the default replacement because we have already handled it.
            return ReplaceAction.Skip;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the existing DOCX document.
            Document doc = new Document("Input.docx");

            // Set up find/replace options with the custom callback.
            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new ReplaceWithParagraphHandler();

            // Perform the find/replace. The replacement string is empty because
            // the callback handles the insertion of the new paragraph.
            doc.Range.Replace(new Regex(@"\[PLACEHOLDER\]"), string.Empty, options);

            // Save the modified document.
            doc.Save("Output.docx");
        }
    }
}
