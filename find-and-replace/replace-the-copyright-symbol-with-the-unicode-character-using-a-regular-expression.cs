using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working directory.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create a sample document containing the copyright placeholder "(c)".
        string inputPath = Path.Combine(workDir, "input.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("All rights reserved (c) 2023.");
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches the literal "(c)".
        Regex copyrightRegex = new Regex(@"\(c\)", RegexOptions.IgnoreCase);

        // Perform the replacement with the Unicode © character.
        int replacedCount = loaded.Range.Replace(copyrightRegex, "©", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one copyright replacement.");

        // Save the modified document.
        string outputPath = Path.Combine(workDir, "output.docx");
        loaded.Save(outputPath);
    }
}
