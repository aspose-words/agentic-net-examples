using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindAndReplaceExample
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Simple text replace: replace all occurrences of "oldText" with "newText".
        doc.Range.Replace("oldText", "newText");

        // Regex replace: replace any word "Sample" followed by digits with "Example".
        FindReplaceOptions options = new FindReplaceOptions();
        // Example of using a callback to log each replacement.
        options.ReplacingCallback = new LogReplacementCallback();

        doc.Range.Replace(new Regex(@"Sample\d+"), "Example", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Callback that logs each replacement and demonstrates returning ReplaceAction.Replace.
    private class LogReplacementCallback : IReplacingCallback
    {
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            Console.WriteLine($"Replacing \"{args.Match.Value}\" with \"{args.Replacement}\" at offset {args.MatchOffset}.");
            // The replacement text is already set in args.Replacement, so just proceed.
            return ReplaceAction.Replace;
        }
    }
}
