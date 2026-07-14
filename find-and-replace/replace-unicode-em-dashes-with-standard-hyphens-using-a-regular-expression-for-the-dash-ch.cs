using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing Unicode em dashes (U+2014).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is an example—text with em dashes—used for replacement.");
        builder.Writeln("Another line—showing multiple—em dashes.");
        // Save the original document (optional, just for inspection).
        doc.Save("input.docx");

        // Load the document to perform find-and-replace.
        Document loaded = new Document("input.docx");

        // Define a regular expression that matches the em dash character.
        Regex emDashRegex = new Regex("\u2014"); // Unicode em dash

        // Replace each em dash with a standard hyphen.
        int replacedCount = loaded.Range.Replace(emDashRegex, "-", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No em dash characters were replaced.");

        // Save the modified document.
        loaded.Save("output.docx");
    }
}
