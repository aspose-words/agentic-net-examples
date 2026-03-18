using System;
using Aspose.Words;
using Aspose.Words.Fields;

class DiscountExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a label for the discount value.
        builder.Writeln("Discount:");

        // Insert an IF field that will evaluate the condition: total > 500.
        // The left side of the comparison is the total amount (hard‑coded here as 600).
        // The right side is the threshold 500.
        // The comparison operator is ">".
        FieldIf discountIf = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
        discountIf.LeftExpression = "600";
        discountIf.ComparisonOperator = ">";
        discountIf.RightExpression = "500";

        // Move the builder to the IF field's separator so we can insert the true/false parts.
        builder.MoveTo(discountIf.Separator);

        // ----- True part (total > 500) -----
        // Insert a nested formula field that calculates total * 0.1.
        // This field will be displayed when the condition is true.
        builder.InsertField("= 600 * 0.1", null);

        // ----- False part (total <= 500) -----
        // After the true part, write the false result – a literal zero.
        builder.Write("0");

        // Update all fields so that their results are calculated.
        doc.UpdateFields();

        // Save the document.
        doc.Save("Discount.docx");
    }
}
