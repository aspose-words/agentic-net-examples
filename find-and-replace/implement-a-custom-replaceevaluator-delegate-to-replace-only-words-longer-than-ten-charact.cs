using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text containing words of various lengths.
        builder.Writeln("This document contains someverylongwordindeed and short words.");
        builder.Writeln("AnotherSuperLongWord is also present.");
        builder.Writeln("Normal words like example and test.");

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions(new LongWordReplacer());

        // Use a regex that matches every word. The callback will decide whether to replace it.
        int replacedCount = doc.Range.Replace(new Regex(@"\b\w+\b"), string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No words longer than ten characters were found.");

        // Save the modified document.
        const string outputPath = "Modified.docx";
        doc.Save(outputPath);

        // Output the result count (no interactive input required).
        Console.WriteLine($"Replaced {replacedCount} word(s). Document saved to '{outputPath}'.");
    }

    // Custom callback that replaces only words longer than ten characters.
    private class LongWordReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // If the matched word length exceeds ten characters, replace it.
            if (args.Match.Value.Length > 10)
            {
                // Example replacement text; can be customized as needed.
                args.Replacement = "[LONG]";
                return ReplaceAction.Replace;
            }

            // Skip replacement for shorter words.
            return ReplaceAction.Skip;
        }
    }
}
