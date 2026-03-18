using System;
using Aspose.Words;
using Aspose.Words.Fields;

class StockStatusExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a label before the field.
        builder.Write("Stock: ");

        // Insert an IF field that checks if the stock quantity is greater than zero.
        // If true, display "In stock"; otherwise display "Out of stock".
        FieldIf stockField = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
        stockField.LeftExpression = "StockQty";          // left side of the comparison
        stockField.ComparisonOperator = ">";            // comparison operator
        stockField.RightExpression = "0";               // right side of the comparison
        stockField.TrueText = "In stock";               // text when condition is true
        stockField.FalseText = "Out of stock";          // text when condition is false
        stockField.Update();                            // evaluate and display the result

        // Save the document to disk.
        doc.Save("StockStatus.docx");
    }
}
