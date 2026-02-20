using System;
using System.Text.RegularExpressions;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Define the text to find (regular expression pattern).
        string pattern = @"PLACEHOLDER";

        // Create FindReplaceOptions and assign a custom replacing callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new RunReplacingCallback();

        // Perform the find-and-replace operation.
        doc.Range.Replace(new Regex(pattern), string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Custom callback that replaces each match with a new Run node.
    private class RunReplacingCallback : IReplacingCallback
    {
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            // Create a new Run with the desired replacement text.
            Run newRun = new Run(e.MatchNode.Document, "NewRunText");

            // Apply formatting to the new Run (example: bold and blue).
            newRun.Font.Bold = true;
            newRun.Font.Color = Color.Blue;

            // The match node is usually a Run inside a Paragraph (CompositeNode).
            // Insert the new Run after the matched node.
            CompositeNode parent = e.MatchNode.ParentNode as CompositeNode;
            if (parent != null)
            {
                parent.InsertAfter(newRun, e.MatchNode);
                // Remove the original node that contained the placeholder.
                e.MatchNode.Remove();
            }

            // Skip the default replacement to avoid inserting the original replacement string.
            return ReplaceAction.Skip;
        }
    }
}
