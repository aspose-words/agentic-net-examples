using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing the text "(c)".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document. (c) 2023 Example Corp.");

        // Save the sample document to a local file.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        doc.Save(inputPath);

        // Load the document we just created.
        Document loadedDoc = new Document(inputPath);

        // Define a regular expression that matches the copyright symbol written as "(c)" (case‑insensitive).
        Regex copyrightRegex = new Regex(@"\(c\)", RegexOptions.IgnoreCase);

        // Replace all matches with the Unicode © character.
        int replacementCount = loadedDoc.Range.Replace(copyrightRegex, "©", new FindReplaceOptions());

        // Ensure that at least one replacement was performed.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one copyright symbol replacement.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        loadedDoc.Save(outputPath);
    }
}
