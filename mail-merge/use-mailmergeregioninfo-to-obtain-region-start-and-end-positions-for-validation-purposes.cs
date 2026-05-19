using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a mail‑merge region named "MyRegion" with two merge fields.
        // The region is delimited by TableStart and TableEnd fields.
        builder.InsertField(" MERGEFIELD TableStart:MyRegion");
        builder.InsertField(" MERGEFIELD Field1");
        builder.Write(", ");
        builder.InsertField(" MERGEFIELD Field2");
        builder.InsertField(" MERGEFIELD TableEnd:MyRegion");

        // Retrieve the full hierarchy of mail‑merge regions in the document.
        MailMergeRegionInfo hierarchy = doc.MailMerge.GetRegionsHierarchy();

        // The top‑level regions are stored in the Regions collection.
        IList<MailMergeRegionInfo> topRegions = hierarchy.Regions;

        // Iterate through each region and obtain its start and end fields.
        foreach (MailMergeRegionInfo region in topRegions)
        {
            // StartField and EndField give access to the underlying MERGEFIELD objects.
            FieldMergeField startField = region.StartField;
            FieldMergeField endField = region.EndField;

            // Output the region name and the names of its start/end fields.
            Console.WriteLine($"Region name: {region.Name}");
            Console.WriteLine($"  Start field name: {startField.FieldName}");
            Console.WriteLine($"  End field name:   {endField.FieldName}");
        }

        // Save the document to verify that the region was created correctly.
        doc.Save("MailMergeRegionInfoExample.docx");
    }
}
