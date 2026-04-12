using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class FindReplaceDemo
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs containing the word "product" in various cases.
        builder.Writeln("Product name: Product");
        builder.Writeln("product name: product");
        builder.Writeln("productivity is high");
        builder.Writeln("The PRODUCT is great.");

        // Configure find-and-replace options: ignore case and match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = true
        };

        // Perform the replacement.
        int replacementCount = doc.Range.Replace("product", "Item", options);

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("No replacements were made.");

        // Save the modified document.
        const string outputPath = "UpdatedDocument.docx";
        doc.Save(outputPath);

        // Output information to the console.
        Console.WriteLine($"Replacements made: {replacementCount}");
        Console.WriteLine("Updated document text:");
        Console.WriteLine(doc.GetText().Trim());
    }
}
