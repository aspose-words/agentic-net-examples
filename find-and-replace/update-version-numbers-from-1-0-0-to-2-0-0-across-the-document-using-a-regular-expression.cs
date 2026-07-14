using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document containing version numbers "1.0.0".
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Release notes:");
        builder.Writeln("Version 1.0.0 – Initial release.");
        builder.Writeln("Bug fixes are applied in version 1.0.0.");
        builder.Writeln("Upcoming version will be 2.0.0.");
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document and replace all occurrences of "1.0.0"
        //    with "2.0.0" using a regular expression.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);
        Regex versionPattern = new Regex(@"\b1\.0\.0\b"); // matches whole "1.0.0"
        int replacedCount = loaded.Range.Replace(versionPattern, "2.0.0", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one version number replacement.");

        // -----------------------------------------------------------------
        // 3. Save the modified document.
        // -----------------------------------------------------------------
        loaded.Save(outputPath);
    }
}
