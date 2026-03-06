using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Create the original document.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);

        // Add body content.
        builder.Writeln("Original body paragraph.");

        // Add a header with some formatted text.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Font.Name = "Arial";
        builder.Font.Size = 12;
        builder.Writeln("Original header text.");

        // Clone the document to create an edited version.
        Document docEdited = (Document)docOriginal.Clone(true);

        // Modify the header formatting in the edited document.
        HeaderFooter editedHeader = docEdited.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (editedHeader != null && editedHeader.FirstParagraph != null && editedHeader.FirstParagraph.Runs.Count > 0)
        {
            // Increase the font size of the first run in the header.
            editedHeader.FirstParagraph.Runs[0].Font.Size = 20;
        }

        // Configure comparison options to ignore header/footer changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreHeadersAndFooters = true
        };

        // Compare the original document with the edited one.
        docOriginal.Compare(docEdited, "Reviewer", DateTime.Now, compareOptions);

        // Save the comparison result.
        docOriginal.Save("ComparisonResult.docx");
    }
}
