using System;
using Aspose.Words;
using Aspose.Words.Fields;

class DiscountInfoExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("Order details:");

        // Insert an IF field that displays the discount only when the Discount merge field value is greater than zero.
        // Field code: IF  { MERGEFIELD Discount }  > 0  "Discount: { MERGEFIELD Discount }%"  ""
        builder.InsertField("IF  { MERGEFIELD Discount }  > 0  \"Discount: { MERGEFIELD Discount }%\"  \"\"");

        // Save the document.
        doc.Save("DiscountInfo.docx");
    }
}
