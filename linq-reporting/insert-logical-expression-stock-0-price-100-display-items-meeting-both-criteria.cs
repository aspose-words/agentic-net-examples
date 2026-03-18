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

        // Insert an outer IF field that checks if the stock value is greater than 0.
        // The field code will be: IF  stock > 0  (true part)  (false part)
        FieldIf outerIf = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
        outerIf.LeftExpression = "stock";
        outerIf.ComparisonOperator = ">";
        outerIf.RightExpression = "0";

        // Prepare the true and false texts for the outer IF.
        // We'll leave the false text empty (nothing displayed when stock <= 0).
        outerIf.FalseText = "";

        // Move the builder to the separator of the outer IF field so we can insert the inner IF.
        builder.MoveTo(outerIf.Separator);

        // Insert an inner IF field that checks if the price is less than 100.
        // This field will be the true part of the outer IF.
        FieldIf innerIf = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
        innerIf.LeftExpression = "price";
        innerIf.ComparisonOperator = "<";
        innerIf.RightExpression = "100";

        // Set the true text of the inner IF to the message we want to display.
        innerIf.TrueText = "Item meets both stock and price criteria.";
        // Leave the false text empty – nothing is shown when price >= 100.
        innerIf.FalseText = "";

        // After inserting the inner IF, we need to close the outer IF's true part.
        // Move the builder to the end of the outer IF field and set its false text (already empty).
        builder.MoveTo(outerIf.End);
        outerIf.TrueText = ""; // The true text is already represented by the inner IF field.

        // Update all fields so that they evaluate based on the current document data.
        doc.UpdateFields();

        // Save the document.
        doc.Save("LogicalExpressionResult.docx");
    }
}
