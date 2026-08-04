using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;   // Needed for FieldMergeField

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a mail merge region named "MyRegion" with two fields.
        builder.InsertField(" MERGEFIELD TableStart:MyRegion");
        builder.InsertField(" MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField(" MERGEFIELD LastName");
        builder.InsertField(" MERGEFIELD TableEnd:MyRegion");

        // Retrieve the full hierarchy of mail merge regions.
        MailMergeRegionInfo hierarchy = doc.MailMerge.GetRegionsHierarchy();

        // Get the top‑level regions (there should be only one in this example).
        IList<MailMergeRegionInfo> topRegions = hierarchy.Regions;

        foreach (MailMergeRegionInfo region in topRegions)
        {
            // Output basic information about the region.
            Console.WriteLine($"Region Name: {region.Name}");
            Console.WriteLine($"Nesting Level: {region.Level}");

            // Obtain the start and end fields for the region.
            FieldMergeField startField = region.StartField;
            FieldMergeField endField = region.EndField;

            // Output the field names.
            Console.WriteLine($"Start Field Name: {startField.FieldName}");
            Console.WriteLine($"End Field Name: {endField.FieldName}");

            // Simple validation of the start/end tags.
            bool startValid = startField.FieldName.StartsWith("TableStart:");
            bool endValid = endField.FieldName.StartsWith("TableEnd:");
            Console.WriteLine($"Start Tag Valid: {startValid}");
            Console.WriteLine($"End Tag Valid: {endValid}");
        }

        // Save the document to verify that the region was created correctly.
        doc.Save("MailMergeRegionInfoExample.docx");
    }
}
