using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and add sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Supercalifragilisticexpialidocious is a famously long word.");
        builder.Writeln("Anotherlongwordexample demonstrates the replacement logic.");
        builder.Writeln("Short words stay untouched.");

        // Regex that matches words longer than ten characters.
        Regex longWordRegex = new Regex(@"\b\w{11,}\b", RegexOptions.Compiled);

        // Set up replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new LongWordReplacer()
        };

        // Perform the replace operation. The replacement string argument is ignored
        // because the callback supplies the actual replacement.
        int replacedCount = doc.Range.Replace(longWordRegex, string.Empty, options);

        // Ensure that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);
    }

    // Callback that replaces each matched long word with "SHORT".
    private class LongWordReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            args.Replacement = "SHORT";
            return ReplaceAction.Replace;
        }
    }
}
