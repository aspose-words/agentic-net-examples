using System;
using System.Globalization;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class TitleCaseReplacer : IReplacingCallback
{
    // Called for each regex match found during the replace operation.
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Convert the matched uppercase word to title case (first letter upper, rest lower).
        string lower = args.Match.Value.ToLowerInvariant();
        string title = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(lower);
        args.Replacement = title; // Set the replacement text.
        return ReplaceAction.Replace;
    }
}

class Program
{
    static void Main()
    {
        // Create a new document and add some sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("THIS is a TEST. ANOTHER LINE WITH UPPERCASE WORDS LIKE ASP.NET and CSharp.");

        // Define a regex that matches whole words consisting of only uppercase letters (2 or more characters).
        Regex upperCaseWordPattern = new Regex(@"\b[A-Z]{2,}\b");

        // Set up find/replace options with our custom callback.
        FindReplaceOptions options = new FindReplaceOptions(new TitleCaseReplacer());

        // Perform the replace operation on the whole document range.
        doc.Range.Replace(upperCaseWordPattern, string.Empty, options);

        // Save the modified document.
        doc.Save("TitleCaseResult.docx");
    }
}
