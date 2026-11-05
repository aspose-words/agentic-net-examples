using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing Unicode em dashes (U+2014).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is an example—text with an em dash.");
        builder.Writeln("Another line—another dash.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the file system.
        Document loaded = new Document(inputPath);

        // Regular expression that matches the em dash character.
        Regex emDashRegex = new Regex("\u2014");

        // Replace each em dash with a standard hyphen.
        int replacedCount = loaded.Range.Replace(emDashRegex, "-", new FindReplaceOptions());

        // Ensure that at least one replacement was performed.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("No em dash characters were replaced.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
