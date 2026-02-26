using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Define the text to be replaced (e.g., a placeholder like [PARA]).
        const string placeholder = "[PARA]";

        // Set up FindReplaceOptions with a custom callback that will replace the placeholder
        // with an entire new paragraph.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertParagraphHandler("This is the inserted paragraph.")
        };

        // Perform the find-and-replace operation.
        doc.Range.Replace(placeholder, string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}

// Custom callback that inserts a new paragraph in place of the matched text.
class InsertParagraphHandler : IReplacingCallback
{
    private readonly string _newParagraphText;

    public InsertParagraphHandler(string newParagraphText)
    {
        _newParagraphText = newParagraphText;
    }

    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // The node that contains the beginning of the match is a Run.
        // Its parent is the Paragraph that holds the placeholder.
        Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;

        // Create a new paragraph with the desired text.
        Paragraph newParagraph = new Paragraph(placeholderParagraph.Document);
        Run run = new Run(placeholderParagraph.Document, _newParagraphText);
        newParagraph.AppendChild(run);

        // Insert the new paragraph after the placeholder paragraph.
        CompositeNode parent = placeholderParagraph.ParentNode;
        parent.InsertAfter(newParagraph, placeholderParagraph);

        // Remove the original paragraph that contained the placeholder.
        placeholderParagraph.Remove();

        // Skip the default replacement because we have already handled it.
        return ReplaceAction.Skip;
    }
}
