using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;

class InsertTableAtReplaceHandler : IReplacingCallback
{
    // Called for each match found during the replace operation.
    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // The node that contains the start of the match is a Run inside a Paragraph.
        // Its parent is the Paragraph that we want to replace.
        Paragraph paragraph = (Paragraph)args.MatchNode.ParentNode;

        // Create a DocumentBuilder attached to the same document.
        // Cast to Document because the overload that accepts DocumentBase was removed in newer versions.
        DocumentBuilder builder = new DocumentBuilder((Document)paragraph.Document);
        // Move the cursor to the paragraph that contains the match.
        builder.MoveTo(paragraph);

        // Build a simple 2‑cell table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Remove the original paragraph that held the placeholder text.
        paragraph.Remove();

        // Skip the default replacement because we have already handled it.
        return ReplaceAction.Skip;
    }
}

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Set up find/replace options with our custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertTableAtReplaceHandler();

        // Replace the placeholder "[TABLE]" with the table created in the callback.
        doc.Range.Replace(new Regex(@"\[TABLE\]"), string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
