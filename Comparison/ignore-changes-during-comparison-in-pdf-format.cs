using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Fields; // Added for FieldDate

class CompareIgnoreFormattingToPdf
{
    static void Main()
    {
        // Create the original document and add some content.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("Hello World! This is the original paragraph.");
        builder.InsertField(" DATE ");

        // Clone the original document to create an edited version.
        Document docEdited = (Document)docOriginal.Clone(true);
        // Make some changes: modify text and change the date field format.
        Paragraph firstParagraph = docEdited.FirstSection.Body.FirstParagraph;
        firstParagraph.Runs[0].Text = "Hello World! This is the edited paragraph.";
        // Cast the first field to FieldDate and change its property.
        FieldDate dateField = docEdited.Range.Fields[0] as FieldDate;
        if (dateField != null)
        {
            dateField.UseLunarCalendar = true;
        }

        // Set comparison options to ignore formatting changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true,          // Ignore formatting differences.
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            CompareMoves = false,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Reviewer", DateTime.Now, compareOptions);

        // Save the comparison result as a PDF file.
        docOriginal.Save("ComparisonResult.pdf", SaveFormat.Pdf);
    }
}
