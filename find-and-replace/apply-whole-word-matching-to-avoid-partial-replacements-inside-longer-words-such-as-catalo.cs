using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add text that contains the target word as a whole word and as part of a longer word.
        builder.Writeln("The catalogue is ready.");
        builder.Writeln("Please refer to the catalogue123 for details.");

        // Configure find‑replace to match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            FindWholeWordsOnly = true
        };

        // Replace the whole word "catalogue" with "catalog".
        int replacementCount = doc.Range.Replace("catalogue", "catalog", options);

        // Ensure that at least one replacement was performed.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document to the local file system.
        const string outputFile = "Modified.docx";
        doc.Save(outputFile);
    }
}
