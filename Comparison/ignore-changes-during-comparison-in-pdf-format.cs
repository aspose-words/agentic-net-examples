using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Fields; // Added for FieldDate

class Program
{
    static void Main()
    {
        // Create the original document and add some content.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("Hello World!");               // Simple paragraph.
        builder.InsertField(" DATE ");                 // Insert a DATE field.

        // Clone the original document to simulate an edited version.
        Document docEdited = (Document)docOriginal.Clone(true);
        // Change the paragraph text.
        docEdited.FirstSection.Body.FirstParagraph.Runs[0].Text = "Hello Aspose!";
        // Modify the DATE field (e.g., use lunar calendar).
        ((FieldDate)docEdited.Range.Fields[0]).UseLunarCalendar = true;

        // Configure comparison options to ignore various changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true,          // Ignore formatting differences.
            IgnoreCaseChanges = true,         // Ignore case changes.
            IgnoreFields = true,              // Ignore field changes.
            // Additional ignore flags can be set here if needed.
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);

        // Save the comparison result as a PDF file.
        docOriginal.Save("ComparisonResult.pdf");
    }
}
