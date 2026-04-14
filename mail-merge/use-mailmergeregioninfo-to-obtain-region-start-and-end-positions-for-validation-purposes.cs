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

        // Insert a mail merge region named "MyRegion" with two merge fields.
        builder.InsertField(" MERGEFIELD TableStart:MyRegion");
        builder.InsertField(" MERGEFIELD Field1");
        builder.Write(", ");
        builder.InsertField(" MERGEFIELD Field2");
        builder.InsertField(" MERGEFIELD TableEnd:MyRegion");

        // Retrieve the full hierarchy of mail merge regions.
        MailMergeRegionInfo hierarchy = doc.MailMerge.GetRegionsHierarchy();

        // Get the list of top‑level regions.
        IList<MailMergeRegionInfo> topRegions = hierarchy.Regions;

        if (topRegions.Count > 0)
        {
            // Work with the first region.
            MailMergeRegionInfo region = topRegions[0];

            // Obtain the start and end fields of the region.
            FieldMergeField startField = region.StartField;
            FieldMergeField endField = region.EndField;

            // Output validation information.
            Console.WriteLine($"Region name: {region.Name}");
            Console.WriteLine($"Nesting level: {region.Level}");
            Console.WriteLine($"Start field name: {startField?.FieldName}");
            Console.WriteLine($"End field name: {endField?.FieldName}");
        }
        else
        {
            Console.WriteLine("No mail merge regions were found in the document.");
        }

        // Save the document (optional, demonstrates the save lifecycle).
        doc.Save("MailMergeRegionInfoExample.docx");
    }
}
