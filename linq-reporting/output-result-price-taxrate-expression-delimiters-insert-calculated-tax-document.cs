using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Sample data – in a real scenario these could come from elsewhere
        double price = 199.99;
        double taxRate = 0.08; // 8%

        // Create a new document and a builder to add content
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write the values with delimiters so they are visible in the document
        builder.Write($"Price: ${price} ");
        builder.Write($"Tax Rate: {taxRate:P0} ");

        // Insert a formula field that calculates price * taxRate
        // The field code will be: = 199.99 * 0.08
        FieldBuilder formulaBuilder = new FieldBuilder(FieldType.FieldFormula);
        formulaBuilder.AddArgument(price);   // left operand
        formulaBuilder.AddArgument("*");     // multiplication operator
        formulaBuilder.AddArgument(taxRate); // right operand

        // Build and insert the field at the current paragraph
        formulaBuilder.BuildAndInsert(builder.CurrentParagraph);

        // Add a label after the field to clarify the result
        builder.Write(" Tax Amount");

        // Recalculate all fields so the formula result is displayed
        doc.UpdateFields();

        // Save the document
        doc.Save("CalculatedTax.docx");
    }
}
