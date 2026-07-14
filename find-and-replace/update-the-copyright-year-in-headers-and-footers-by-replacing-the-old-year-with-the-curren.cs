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
        // Create a sample document with a header and a footer containing an old year.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("(C) 2022 My Company");

        // Insert footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("(C) 2022 My Company");

        // Save the sample document.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document and replace the old year with the current year
        // in all headers and footers.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        int currentYear = DateTime.Now.Year;
        string replacementText = $"(C) {currentYear}";

        // Regex to find the pattern "(C) 4-digit-year".
        Regex yearPattern = new Regex(@"\(C\) \d{4}");

        int totalReplacements = 0;

        // Iterate through every section and its header/footer collection.
        foreach (Section section in loaded.Sections)
        {
            foreach (HeaderFooter headerFooter in section.HeadersFooters)
            {
                // Perform the replacement within the header/footer range.
                int replaced = headerFooter.Range.Replace(yearPattern, replacementText, new FindReplaceOptions());
                totalReplacements += replaced;
            }
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one replacement in headers or footers.");

        // Save the modified document.
        loaded.Save(outputPath);

        // Optional: indicate success.
        Console.WriteLine($"Replaced {totalReplacements} occurrence(s). Output saved to '{outputPath}'.");
    }
}
