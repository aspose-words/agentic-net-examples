using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

class UpdateTableCaptions
{
    static void Main()
    {
        // Load an existing document (or create a new one if you prefer).
        Document doc = new Document("Input.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Example: add a new table to the document.
        // -----------------------------------------------------------------
        builder.Writeln();                     // start a new paragraph before the table
        Table table = builder.StartTable();    // begin table construction

        // First row
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.EndTable();                    // finish the table

        // Insert a caption for the newly added table.
        // The caption consists of a SEQ field with the identifier "Table".
        builder.MoveTo(table.FirstRow.FirstCell.FirstParagraph); // position before the table
        builder.InsertParagraph();            // create a paragraph that will hold the caption
        builder.MoveTo(builder.CurrentParagraph);
        Field seqField = builder.InsertField(FieldType.FieldSequence, true);
        ((FieldSeq)seqField).SequenceIdentifier = "Table";
        builder.Write(" Table Caption");       // optional descriptive text

        // -----------------------------------------------------------------
        // Refresh all table captions in the document.
        // By updating all fields, the SEQ fields that form the captions
        // will be recalculated and display correct sequential numbers.
        // -----------------------------------------------------------------
        doc.UpdateFields();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
