using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FractionQuantityExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document and a builder to work with it.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Example quantity that we want to display as a fraction.
            double quantity = 2.75; // 2 ¾

            // Insert a formula field that evaluates the quantity and applies a fraction number format.
            // The "\\#" switch specifies a numeric format; "# ?/?" tells Aspose.Words to render the result as a fraction.
            builder.InsertField($"= {quantity} \\# \"# ?/?\"");

            // Update the field so that the result is calculated and formatted.
            doc.Range.Fields[0].Update();

            // Save the document (adjust the path as needed).
            doc.Save(@"C:\Temp\FractionQuantity.docx");
        }
    }
}
