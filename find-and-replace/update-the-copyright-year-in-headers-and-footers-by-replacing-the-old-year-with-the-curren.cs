using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder and file paths.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string inputPath = Path.Combine(artifactsDir, "input.docx");
        string outputPath = Path.Combine(artifactsDir, "output.docx");

        // -----------------------------------------------------------------
        // Create a sample document that contains a header and a footer with
        // an old copyright year (2022). This document will be saved and then
        // processed.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Header with old year.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("(C) 2022 Aspose Pty Ltd.");

        // Footer with old year.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("(C) 2022 Aspose Pty Ltd.");

        // Add a simple body paragraph so the document is not empty.
        builder.MoveToDocumentEnd();
        builder.Writeln("Sample body text.");

        // Save the source document.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document and replace the old year with the current year
        // in all headers and footers.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        int currentYear = DateTime.Now.Year;
        string oldYear = "2022";
        string newYear = currentYear.ToString();

        int totalReplacements = 0;
        FindReplaceOptions replaceOptions = new FindReplaceOptions();

        foreach (Section section in loaded.Sections)
        {
            // Process primary header.
            HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header != null)
            {
                int replaced = header.Range.Replace(oldYear, newYear, replaceOptions);
                totalReplacements += replaced;
            }

            // Process primary footer.
            HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            if (footer != null)
            {
                int replaced = footer.Range.Replace(oldYear, newYear, replaceOptions);
                totalReplacements += replaced;
            }
        }

        // Ensure that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one replacement in headers or footers.");

        // Save the modified document.
        loaded.Save(outputPath);
    }
}
