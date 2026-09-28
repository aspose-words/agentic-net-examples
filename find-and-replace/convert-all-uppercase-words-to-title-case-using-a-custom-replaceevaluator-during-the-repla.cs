using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample text containing uppercase words.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("THIS is a TEST document.");
        builder.Writeln("IT contains SOME UPPERCASE WORDS like EXAMPLE and SAMPLE.");

        // Define a regex that matches whole words composed only of uppercase letters.
        Regex uppercaseWordRegex = new Regex(@"\b[A-Z]+\b");

        // Set up find‑replace options with a custom callback that converts each match to title case.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new UppercaseToTitleCaseCallback();

        // Perform the replacement using the overload that accepts a regex, a dummy replacement string,
        // and the FindReplaceOptions containing the callback.
        int replacedCount = doc.Range.Replace(uppercaseWordRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No uppercase words were replaced.");

        // Save the modified document.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }

    // Callback that receives each regex match and replaces it with its title‑cased version.
    private class UppercaseToTitleCaseCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            string word = args.Match.Value;
            if (!string.IsNullOrEmpty(word))
            {
                // Convert first character to upper case and the rest to lower case.
                string titleCase = char.ToUpper(word[0]) + word.Substring(1).ToLower();
                args.Replacement = titleCase;
            }
            return ReplaceAction.Replace;
        }
    }
}
