using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a header with an old copyright year.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("(C) 2022 Aspose Pty Ltd.");

        // Add a footer with the same old copyright year.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("(C) 2022 Aspose Pty Ltd.");

        // Save the original document (optional, demonstrates creation).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Prepare the replacement: current year as a string.
        string currentYear = DateTime.Now.Year.ToString();

        // Use a regular expression to find any four‑digit year.
        Regex yearRegex = new Regex(@"\b\d{4}\b");

        // Configure find‑replace options (case‑insensitive, whole‑word not required).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        int totalReplacements = 0;

        // Iterate through all sections and their headers/footers.
        foreach (Section section in doc.Sections)
        {
            foreach (HeaderFooter headerFooter in section.HeadersFooters)
            {
                // Perform the replacement within the header/footer range.
                int replaced = headerFooter.Range.Replace(yearRegex, currentYear, options);
                totalReplacements += replaced;
            }
        }

        // Validate that at least one replacement was made.
        if (totalReplacements == 0)
            throw new InvalidOperationException("No year was replaced in headers or footers.");

        // Save the updated document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
