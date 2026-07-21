using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document contains someverylongwordindeed and anotherextremelylengthyterm.");
        builder.Writeln("Short words stay unchanged.");

        // Regex that matches words longer than ten characters.
        Regex longWordRegex = new Regex(@"\b\w{11,}\b");

        // Set up replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new LongWordReplacer()
        };

        // Perform the replace. The callback supplies the replacement text.
        int replacedCount = doc.Range.Replace(longWordRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that replaces each matched long word with a placeholder.
    private class LongWordReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            args.Replacement = "[REPLACED]";
            return ReplaceAction.Replace;
        }
    }
}
