using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text.
        builder.Writeln("Order Summary:");

        // Insert an IF field that will display a discount line only when the discount amount is greater than 0.
        // The condition will be based on a MERGEFIELD named \"Discount\".
        // Field code format: IF { MERGEFIELD Discount } > 0 "Discount Applied: $<<Discount>>" ""
        // First, insert the IF field.
        FieldIf ifField = (FieldIf)builder.InsertField(FieldType.FieldIf, true);

        // Build the left expression: a MERGEFIELD that will be replaced during mail merge.
        // Move the cursor to the start of the IF field to insert the MERGEFIELD.
        builder.MoveTo(ifField.Start);
        builder.InsertField(" MERGEFIELD Discount ");

        // Set the left expression to the field code we just inserted.
        // The DocumentBuilder automatically places the MERGEFIELD inside the IF field code.
        // Now configure the rest of the IF field.
        ifField.ComparisonOperator = ">";
        ifField.RightExpression = "0";

        // Text to display when the condition is true (discount exists).
        ifField.TrueText = "Discount Applied: $<<Discount>>";

        // Text to display when the condition is false (no discount).
        ifField.FalseText = "";

        // Update the field so that the result is calculated (useful for preview).
        ifField.Update();

        // Add a line break after the IF field.
        builder.Writeln();

        // Add other content to the document.
        builder.Writeln("Thank you for your purchase!");

        // Save the document.
        doc.Save("ConditionalDiscount.docx");
    }
}
