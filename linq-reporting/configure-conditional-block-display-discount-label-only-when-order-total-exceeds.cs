using System;
using Aspose.Words;
using Aspose.Words.Fields;

class DiscountLabelExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading for the order.
        builder.Writeln("Order Summary:");

        // Insert an IF field that shows a discount label only when Order.Total > 100.
        // The field will evaluate the expression: Order.Total > 100
        // If true, it displays "Discount: 10%"; otherwise it displays nothing.
        FieldIf discountField = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
        discountField.LeftExpression = "Order.Total";      // Left side of the comparison.
        discountField.ComparisonOperator = ">";           // Comparison operator.
        discountField.RightExpression = "100";            // Right side (threshold).
        discountField.TrueText = "Discount: 10%";         // Text shown when condition is true.
        discountField.FalseText = "";                     // Text shown when condition is false.
        discountField.Update();                           // Calculate the field result.

        // Add a line break after the field.
        builder.Writeln();

        // Save the document to disk.
        doc.Save("DiscountLabel.docx");
    }
}
