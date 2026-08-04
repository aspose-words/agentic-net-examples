using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;      // Required package reference
using Newtonsoft.Json;    // Required package reference

public class Program
{
    public static void Main()
    {
        // Define file names for the sample input and output documents.
        const string inputFile = "input.docx";
        const string outputFile = "output.docx";

        // -----------------------------------------------------------------
        // Create a sample document containing version numbers to replace.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Current release: 1.0.0");
        builder.Writeln("Previous release: 1.0.0");
        builder.Writeln("No change needed for 2.1.3");
        doc.Save(inputFile);

        // -----------------------------------------------------------------
        // Load the document and perform a regex replace from 1.0.0 to 2.0.0.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputFile);

        // Regex that matches the exact version string "1.0.0".
        Regex versionPattern = new Regex(@"\b1\.0\.0\b");

        // Use FindReplaceOptions as required by the API.
        FindReplaceOptions options = new FindReplaceOptions();

        int replacedCount = loaded.Range.Replace(versionPattern, "2.0.0", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one version number replacement.");

        // Save the modified document.
        loaded.Save(outputFile);
    }
}
