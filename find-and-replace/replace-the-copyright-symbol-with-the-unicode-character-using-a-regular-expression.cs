using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words for drawing types
using Newtonsoft.Json; // Included as per package requirement

public class Program
{
    public static void Main()
    {
        // Create a new document and add a paragraph containing the (c) symbol.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample (c) text. Another (c) occurrence.");

        // Define a regular expression that matches the literal "(c)" string, case‑insensitive.
        Regex regex = new Regex(@"\(c\)", RegexOptions.IgnoreCase);

        // Perform the replacement with the Unicode copyright character.
        int replacedCount = doc.Range.Replace(regex, "©", new FindReplaceOptions());

        // Validate that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were performed.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Optionally, confirm the file was created (throws if not).
        if (!File.Exists(outputPath))
            throw new FileNotFoundException($"The output file '{outputPath}' was not created.");

        // End of example – no interactive prompts.
    }
}
