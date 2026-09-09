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
        builder.Writeln("This example contains supercalifragilisticexpialidocious and anotherlongword and short.");

        // Regular expression that matches whole words.
        Regex wordRegex = new Regex(@"\b\w+\b", RegexOptions.Compiled);

        // Set up find/replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new LongWordReplacer();

        // Perform the replace operation using the callback.
        int replacements = doc.Range.Replace(wordRegex, string.Empty, options);

        // Verify that at least one replacement was performed.
        if (replacements == 0)
            throw new InvalidOperationException("Expected at least one replacement for long words.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Ensure the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }

    // Callback that replaces words longer than 10 characters with "SHORT".
    private class LongWordReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // If the matched word is longer than 10 characters, replace it.
            if (args.Match.Value.Length > 10)
            {
                args.Replacement = "SHORT";
                return ReplaceAction.Replace;
            }

            // Otherwise, keep the original word unchanged.
            return ReplaceAction.Skip;
        }
    }
}
