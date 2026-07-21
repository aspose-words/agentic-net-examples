using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by the category rules
using Newtonsoft.Json; // Required by the category rules

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text containing Unicode em dashes (U+2014).
        builder.Writeln("This is an example—text with em dashes—used for testing.");
        builder.Writeln("Another line—showing multiple—em dashes.");

        // Save the original document (optional, for inspection).
        doc.Save("input.docx");

        // Define a regular expression that matches the em dash character.
        Regex emDashRegex = new Regex("\u2014");

        // Perform the replacement: em dash → hyphen.
        int replacedCount = doc.Range.Replace(emDashRegex, "-", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one em dash replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }
}
