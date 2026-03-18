using System;
using Aspose.Words;
using Aspose.Words.Fields;

class ConditionalStockExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text.
        builder.Writeln("Product availability:");

        // Insert an IF field that evaluates the MERGEFIELD $stock$.
        // The field will display "In stock" when $stock$ > 0, otherwise nothing.
        FieldIf ifField = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
        ifField.LeftExpression = "$stock$";          // MERGEFIELD placeholder.
        ifField.ComparisonOperator = ">";           // Greater‑than operator.
        ifField.RightExpression = "0";              // Compare against zero.
        ifField.TrueText = "In stock";              // Text shown when condition is true.
        ifField.FalseText = "";                     // Text shown when condition is false.
        ifField.Update();                           // Calculate the result immediately.

        // Add a line break after the field for readability.
        builder.Writeln();

        // Save the document.
        doc.Save("ConditionalStock.docx");
    }
}
