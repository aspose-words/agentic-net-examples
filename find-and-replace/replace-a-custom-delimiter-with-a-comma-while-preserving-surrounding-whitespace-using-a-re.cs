using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create a sample document containing a custom delimiter (|).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Apple | Banana|Cherry | Date");
        // Save the document so it can be loaded later.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document and replace the custom delimiter with a comma,
        // preserving any surrounding whitespace.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // Regex captures optional whitespace before and after the delimiter.
        Regex delimiterPattern = new Regex(@"(\s*)\|(\s*)");
        // Replacement keeps the captured whitespace groups and inserts a comma.
        const string replacement = "$1,$2";

        // Perform the replacement using Aspose.Words' Range.Replace method.
        int replacementCount = loaded.Range.Replace(delimiterPattern, replacement, new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one delimiter replacement.");

        // Save the modified document.
        loaded.Save(outputPath);
    }
}
