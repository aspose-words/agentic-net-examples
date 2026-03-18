using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Example product price.
        double price = 1234.56;

        // Insert a formula field that evaluates the price and formats it as currency.
        // The field code uses the numeric format switch (\#) with a standard .NET format string.
        // Field code: = 1234.56 \# "$#,0.00"
        builder.InsertField($"= {price} \\# \"$#,0.00\"");

        // Calculate the field result.
        doc.UpdateFields();

        // Save the document.
        doc.Save("ProductPrice.docx");
    }
}
