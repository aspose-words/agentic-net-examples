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
        builder.Writeln("Sample hyperlinks:");
        builder.InsertHyperlink("http://example.com/page1", "Example 1", false);
        builder.Writeln();
        builder.InsertHyperlink("http://test.org/resource", "Test Resource", false);
        builder.Writeln();
        builder.InsertHyperlink("http://sub.domain.com", "Subdomain", false);
        builder.Writeln();

        // Save the original document (optional, for inspection).
        doc.Save("original.docx");

        // Define a regular expression that matches the http scheme at the start of a URL.
        Regex httpScheme = new Regex(@"http://", RegexOptions.IgnoreCase);

        // Perform the replacement across the whole document.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = doc.Range.Replace(httpScheme, "https://", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No hyperlink URLs were updated.");

        // Save the modified document.
        doc.Save("updated.docx");
    }
}
