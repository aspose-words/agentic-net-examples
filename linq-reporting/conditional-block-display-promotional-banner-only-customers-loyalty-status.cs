using System;
using Aspose.Words;
using Aspose.Words.Fields;

class PromotionalBannerExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an IF field that checks the value of a MERGEFIELD named "LoyaltyStatus".
        // If the status equals "Gold", the true part will display the promotional banner text.
        // Otherwise nothing will be shown.
        FieldIf ifField = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
        // Left side of the comparison – the merge field value.
        ifField.LeftExpression = "\"{ MERGEFIELD LoyaltyStatus }\"";
        // Comparison operator.
        ifField.ComparisonOperator = "=";
        // Right side of the comparison – the loyalty level we care about.
        ifField.RightExpression = "\"Gold\"";
        // Text displayed when the condition is true (the promotional banner).
        ifField.TrueText = "Special Promotion: 20% off your next purchase!";
        // Text displayed when the condition is false – empty string means no banner.
        ifField.FalseText = "";
        // Evaluate the field so the result is stored in the document.
        ifField.Update();

        // Save the document to disk.
        doc.Save("PromotionalBanner.docx");
    }
}
