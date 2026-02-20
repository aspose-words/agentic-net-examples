using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace FindReplaceDemo
{
    // Custom callback that logs each match and performs the replacement.
    class LoggingReplacingCallback : IReplacingCallback
    {
        private readonly StringBuilder _log = new StringBuilder();

        // This method is called for each match found during the replace operation.
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Log the original matched text and its location.
            _log.AppendLine($"Matched \"{args.Match.Value}\" at offset {args.MatchOffset} in a {args.MatchNode.NodeType} node.");

            // Optionally modify the replacement text.
            // args.Replacement = $"[Found:{args.Match.Value}]";

            // Proceed with the default replacement.
            return ReplaceAction.Replace;
        }

        // Expose the log for external use.
        public string GetLog() => _log.ToString();
    }

    class Program
    {
        static void Main()
        {
            // Load the source DOCX document.
            Document doc = new Document("Input.docx");

            // Define the text pattern to find. Here we look for the word "Hello" (case‑insensitive).
            Regex findPattern = new Regex(@"\bHello\b", RegexOptions.IgnoreCase);

            // Create a FindReplaceOptions instance and assign the custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new LoggingReplacingCallback()
            };

            // Perform the find‑and‑replace operation, replacing matches with "Hi".
            doc.Range.Replace(findPattern, "Hi", options);

            // Save the modified document.
            doc.Save("Output.docx");

            // If needed, retrieve and display the log from the callback.
            var callback = (LoggingReplacingCallback)options.ReplacingCallback;
            Console.WriteLine("Replacement log:");
            Console.WriteLine(callback.GetLog());
        }
    }
}
