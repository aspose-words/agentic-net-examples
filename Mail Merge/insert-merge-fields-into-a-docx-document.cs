using System;
using Aspose.Words;
using Aspose.Words.Fields;

class MergeFieldExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which simplifies inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three MERGEFIELDs into the document.
        // The field code is written without the surrounding braces.
        builder.InsertField(@"MERGEFIELD Chairman \* MERGEFORMAT");
        builder.InsertParagraph(); // Add a paragraph break between fields.
        builder.InsertField(@"MERGEFIELD ChiefFinancialOfficer \* MERGEFORMAT");
        builder.InsertParagraph();
        builder.InsertField(@"MERGEFIELD ChiefTechnologyOfficer \* MERGEFORMAT");

        // Fill the MERGEFIELDs manually using the builder.
        // MoveToMergeField moves the cursor to the field and removes it from the document.
        // After moving, we can write the desired value.
        builder.MoveToMergeField("Chairman");
        builder.Bold = true;               // Example formatting.
        builder.Writeln("John Doe");

        builder.MoveToMergeField("ChiefFinancialOfficer");
        builder.Italic = true;
        builder.Writeln("Jane Doe");

        builder.MoveToMergeField("ChiefTechnologyOfficer");
        builder.Italic = true;
        builder.Writeln("John Bloggs");

        // Save the document to a DOCX file.
        doc.Save("MergeFieldsFilled.docx");
    }
}
