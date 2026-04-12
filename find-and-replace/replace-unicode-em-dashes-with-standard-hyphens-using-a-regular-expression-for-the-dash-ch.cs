using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a paragraph that contains an em dash (U+2014).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is an example—text with an em dash.");

        // Define a regular expression that matches the Unicode em dash character.
        Regex emDashPattern = new Regex("\u2014");

        // Replace all occurrences of the em dash with a standard hyphen.
        int replacementCount = doc.Range.Replace(emDashPattern, "-");

        // Validate that at least one replacement was performed.
        if (replacementCount == 0)
            throw new InvalidOperationException("No em dash characters were found to replace.");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
