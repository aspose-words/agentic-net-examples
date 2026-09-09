using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing the "(c)" copyright placeholder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample text with (c) 2023 Company.");

        // Save the initial document (optional, demonstrates the create‑save workflow).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document to perform find‑and‑replace.
        Document loaded = new Document(inputPath);

        // Regular expression to match the "(c)" pattern (case‑insensitive).
        Regex copyrightPattern = new Regex(@"\(c\)", RegexOptions.IgnoreCase);

        // Replace the pattern with the Unicode © character.
        int replacementCount = loaded.Range.Replace(copyrightPattern, "©", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
