using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing plain‑text URLs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Visit our site: http://example.com/page");
        builder.Writeln("Another link: http://test.org");
        // Save the source document (optional, just to demonstrate the workflow).
        doc.Save("input.docx");

        // Load the document we just created.
        Document loaded = new Document("input.docx");

        // Regular expression that captures the part after "http://".
        Regex httpRegex = new Regex(@"http://([^ \t\r\n]+)", RegexOptions.IgnoreCase);
        // Replacement string keeps the captured part and prefixes it with "https://".
        string replacement = "https://$1";

        // Perform the find‑and‑replace operation.
        int replacedCount = loaded.Range.Replace(httpRegex, replacement, new FindReplaceOptions());

        // Validate that at least one URL was updated.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one hyperlink scheme replacement.");

        // Save the modified document.
        loaded.Save("output.docx");
    }
}
