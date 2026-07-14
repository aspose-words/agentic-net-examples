using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with text that matches a regular expression pattern.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Order 123 and Order 456");
        doc.Save("input.docx");

        // Load the document we just created.
        Document loaded = new Document("input.docx");

        // Perform the regex‑based replacement using the Regex overload.
        int replacedCount = loaded.Range.Replace(new Regex(@"Order \d+"), "Order ###", new FindReplaceOptions());

        // Ensure that the expected number of replacements occurred.
        if (replacedCount != 2)
            throw new InvalidOperationException($"Expected 2 replacements, but got {replacedCount}.");

        // Save the modified document.
        loaded.Save("output.docx");
    }
}
