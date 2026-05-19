using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing the text "(c)" which we will replace.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document. (c) 2023 Example Corp.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the file system.
        Document loadedDoc = new Document(inputPath);

        // Define a regular expression that matches the literal "(c)" (case‑insensitive).
        Regex copyrightPattern = new Regex(@"\(c\)", RegexOptions.IgnoreCase);

        // Replace each match with the Unicode copyright character.
        int replacedCount = loadedDoc.Range.Replace(copyrightPattern, "©", new FindReplaceOptions());

        // Ensure that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one copyright symbol replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
