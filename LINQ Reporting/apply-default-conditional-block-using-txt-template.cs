using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a heading.
        builder.Writeln("Result of conditional block:");

        // Insert an IF field: if 5 = 5 then display "Yes", otherwise "No".
        FieldIf ifField = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
        ifField.LeftExpression = "5";
        ifField.ComparisonOperator = "=";
        ifField.RightExpression = "5";
        ifField.TrueText = "Yes";
        ifField.FalseText = "No";
        ifField.Update();

        // Update all fields before saving.
        doc.UpdateFields();

        // Configure TXT save options – keep the default template.
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        txtOptions.DefaultTemplate = ""; // No external template.
        txtOptions.ParagraphBreak = "\r\n"; // Use standard line break.

        // Save the document as plain text.
        doc.Save("ConditionalBlock.txt", txtOptions);
    }
}
