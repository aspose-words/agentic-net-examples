using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words for drawing types

public class Program
{
    public static void Main()
    {
        // Create a sample document containing a custom delimiter (pipe character).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Apple | Banana | Cherry");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document to perform the find‑and‑replace operation.
        Document loaded = new Document(inputPath);

        // Regex pattern:
        //   (\s*)   – captures any whitespace before the delimiter.
        //   \|      – matches the pipe character (custom delimiter).
        //   (\s*)   – captures any whitespace after the delimiter.
        // The replacement string re‑inserts the captured whitespace and inserts a comma.
        Regex delimiterRegex = new Regex(@"(\s*)\|(\s*)");
        string replacement = "$1,$2";

        // Perform the replacement using Aspose.Words Range.Replace.
        int replacedCount = loaded.Range.Replace(delimiterRegex, replacement, new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one delimiter replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
