using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Simple string replace across the whole document.
        doc.Range.Replace("old value", "new value", new FindReplaceOptions());

        // Regular‑expression replace with a callback to log each replacement.
        FindReplaceOptions options = new FindReplaceOptions(new ReplacementLogger());
        doc.Range.Replace(new Regex(@"\b(old value)\b"), "new value", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Callback that is invoked for every match found during a replace operation.
    private class ReplacementLogger : IReplacingCallback
    {
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            // Log the original text and the replacement that will be applied.
            Console.WriteLine($"Replacing \"{e.Match.Value}\" with \"{e.Replacement}\"");

            // Optionally modify the replacement text here:
            // e.Replacement = $"[{e.Match.Value}]";

            // Continue with the replacement.
            return ReplaceAction.Replace;
        }
    }
}
