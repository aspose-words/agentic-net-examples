using System;
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

        // Insert sample hyperlinks that use the http scheme.
        builder.InsertHyperlink("Aspose", "http://www.aspose.com", false);
        builder.Writeln();
        builder.InsertHyperlink("GitHub", "http://github.com/aspose-words", false);
        builder.Writeln();

        // Define a regular expression that matches the http scheme.
        Regex httpPattern = new Regex(@"http://", RegexOptions.IgnoreCase);

        // Perform the replacement across the whole document.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = doc.Range.Replace(httpPattern, "https://", options);

        // Ensure that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("No http URLs were found to replace.");

        // Update fields so that hyperlink targets reflect the changed URLs.
        doc.UpdateFields();

        // Save the modified document.
        const string outputPath = "HyperlinksUpdated.docx";
        doc.Save(outputPath);

        // Output the result count (optional, not required for the task).
        Console.WriteLine($"Replaced {replacedCount} occurrence(s). Document saved to '{outputPath}'.");
    }
}
