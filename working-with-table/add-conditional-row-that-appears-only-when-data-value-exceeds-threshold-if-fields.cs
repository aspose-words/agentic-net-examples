using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

class ConditionalRowExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a header row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Score");
        builder.EndRow();

        // Insert a regular data row (always shown).
        builder.InsertCell();
        builder.Write("Widget A");
        builder.InsertCell();
        // Merge field that will hold the numeric value.
        builder.InsertField("MERGEFIELD Score \\* MERGEFORMAT");
        builder.EndRow();

        // Insert an IF field that will conditionally display the next row.
        // The IF field compares the value of the MERGEFIELD "Score" with the threshold (80).
        FieldIf ifField = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
        ifField.LeftExpression = "MERGEFIELD Score";
        ifField.ComparisonOperator = ">";
        ifField.RightExpression = "80";
        // The true part will be filled after we move to the field separator.
        ifField.TrueText = string.Empty; // placeholder; actual content will be inserted after the separator.
        ifField.FalseText = string.Empty; // nothing when condition is false.

        // Move the builder to the separator of the IF field.
        // Anything inserted here becomes the "true" output of the IF field.
        builder.MoveTo(ifField.Separator);

        // Insert the conditional row – it will only appear when Score > 80.
        builder.InsertCell();
        builder.Write("Widget A (High Score)");
        builder.InsertCell();
        builder.Write("Excellent");
        builder.EndRow();

        // Close the table.
        builder.EndTable();

        // Prepare a data source for mail merge.
        DataTable data = new DataTable("Data");
        data.Columns.Add("Score");
        // Row where the condition is false (Score <= 80) – the conditional row will be omitted.
        data.Rows.Add("75");
        // Row where the condition is true (Score > 80) – the conditional row will be shown.
        data.Rows.Add("92");

        // Execute mail merge.
        doc.MailMerge.Execute(data);

        // Save the resulting document.
        doc.Save("ConditionalRow.docx");
    }
}
