using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Fields;      // Added for FieldDate
using Aspose.Words.Tables;      // Added for Table

class Program
{
    static void Main()
    {
        // Create the original document and add some content.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("Hello World!");                     // Simple paragraph.
        builder.InsertField(" DATE ");                       // Date field.
        builder.Writeln("This is a sample paragraph.");      // Another paragraph.

        // Add a table with two cells.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndTable();

        // Clone the original document to create the edited version.
        Document docEdited = (Document)docOriginal.Clone(true);

        // Introduce changes in the edited document.
        // Change the text of the first paragraph.
        docEdited.FirstSection.Body.FirstParagraph.Runs[0].Text = "Hello Universe!";

        // Modify the date field to use the lunar calendar.
        ((FieldDate)docEdited.Range.Fields[0]).UseLunarCalendar = true;

        // Change the text inside the second cell of the table.
        ((Table)docEdited.GetChild(NodeType.Table, 0, true))
            .FirstRow.Cells[1].FirstParagraph.Runs[0].Text = "Edited Cell";

        // Configure comparison options to ignore all possible changes.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = false,
            IgnoreFormatting = true,
            IgnoreCaseChanges = true,
            IgnoreComments = true,
            IgnoreTables = true,
            IgnoreFields = true,
            IgnoreFootnotes = true,
            IgnoreTextboxes = true,
            IgnoreHeadersAndFooters = true,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Only non‑ignored differences will generate revisions.
        docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);

        // Save the comparison result as a PDF file.
        docOriginal.Save("ComparisonResult.pdf", SaveFormat.Pdf);
    }
}
