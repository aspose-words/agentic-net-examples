using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a footer that contains placeholder text and page number fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some body content.
        builder.Writeln("This is the main document body.");

        // Move to the primary footer and add placeholder text and page number fields.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Confidential - CompanyName ");
        builder.InsertField(FieldType.FieldPage, true);
        builder.Write(" of ");
        builder.InsertField(FieldType.FieldNumPages, true);

        // Save the input document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Prepare find-and-replace options.
        FindReplaceOptions options = new FindReplaceOptions();

        // Replace the placeholder text "CompanyName" with "NewCompany" in all footers,
        // preserving the page number fields.
        int totalReplacements = 0;
        HeaderFooterType[] footerTypes = new[]
        {
            HeaderFooterType.FooterPrimary,
            HeaderFooterType.FooterFirst,
            HeaderFooterType.FooterEven
        };

        foreach (Section section in loaded.Sections)
        {
            foreach (HeaderFooterType type in footerTypes)
            {
                HeaderFooter footer = section.HeadersFooters[type];
                if (footer != null)
                {
                    int replaced = footer.Range.Replace("CompanyName", "NewCompany", options);
                    totalReplacements += replaced;
                }
            }
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one replacement in footers.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
