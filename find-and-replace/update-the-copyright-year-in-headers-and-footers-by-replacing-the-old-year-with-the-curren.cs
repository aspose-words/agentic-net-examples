using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a header and a footer that contain the old year.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("(C) 2022 My Company");

        // Footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("(C) 2022 My Company");

        // Save the sample document.
        const string inputPath = "sample.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define the old and new year strings.
        string oldYear = "2022";
        string newYear = DateTime.Now.Year.ToString();

        // Options for the replace operation.
        FindReplaceOptions options = new FindReplaceOptions();

        // Perform the replacement in every header and footer of each section.
        int totalReplacements = 0;
        foreach (Section section in loaded.Sections)
        {
            HeaderFooterCollection hfCollection = section.HeadersFooters;
            foreach (HeaderFooter hf in hfCollection)
            {
                int replaced = hf.Range.Replace(oldYear, newYear, options);
                totalReplacements += replaced;
            }
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the updated document.
        const string outputPath = "updated.docx";
        loaded.Save(outputPath);
    }
}
