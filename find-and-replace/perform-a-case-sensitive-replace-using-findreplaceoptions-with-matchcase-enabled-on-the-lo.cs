using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;      // Required package, not used directly in this example
using Newtonsoft.Json;    // Required package, not used directly in this example

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create a sample document with mixed‑case occurrences of the word.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Apple apple APPLE");
        doc.Save(inputPath);

        // ---------------------------------------------------------------
        // Load the document and perform a case‑sensitive replacement.
        // ---------------------------------------------------------------
        Document loaded = new Document(inputPath);
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true   // Enable case‑sensitive matching.
        };

        int replacedCount = loaded.Range.Replace("Apple", "Orange", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        loaded.Save(outputPath);
    }
}
