using System;
using Aspose.Words;
using Aspose.Words.Fields;

class CalculateTotalLineAmount
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph that will hold the calculation.
        builder.Writeln("Total line amount:");

        // Build a MERGEFIELD for the price value.
        FieldBuilder priceField = new FieldBuilder(FieldType.FieldMergeField);
        priceField.AddArgument("price"); // <<[price]>>

        // Build a MERGEFIELD for the quantity value.
        FieldBuilder quantityField = new FieldBuilder(FieldType.FieldMergeField);
        quantityField.AddArgument("quantity"); // <<[quantity]>>

        // Build a FORMULA field that multiplies the two merge fields.
        // The field code will look like:
        // = { MERGEFIELD price } * { MERGEFIELD quantity }
        FieldBuilder formulaField = new FieldBuilder(FieldType.FieldFormula);
        formulaField.AddArgument(priceField);   // first argument: nested MERGEFIELD price
        formulaField.AddArgument("*");         // multiplication operator
        formulaField.AddArgument(quantityField); // second argument: nested MERGEFIELD quantity

        // Insert the formula field at the end of the current paragraph.
        // The result will be the total amount (price * quantity).
        formulaField.BuildAndInsert(builder.CurrentParagraph);

        // Update all fields so the formula is evaluated (useful when the document is opened later).
        doc.UpdateFields();

        // Save the document.
        doc.Save("TotalLineAmount.docx");
    }
}
