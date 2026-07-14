using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;   // Needed for FieldMergeField

public class MailMergeRegionInfoExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a mail merge region named "SampleRegion".
        // The region is defined by TableStart and TableEnd merge fields.
        builder.InsertField(" MERGEFIELD TableStart:SampleRegion");
        builder.InsertField(" MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField(" MERGEFIELD LastName");
        builder.InsertField(" MERGEFIELD TableEnd:SampleRegion");

        // Save the document (optional, just to see the result).
        doc.Save("MailMergeRegionInfoExample.docx");

        // Retrieve the full hierarchy of mail merge regions.
        MailMergeRegionInfo hierarchy = doc.MailMerge.GetRegionsHierarchy();

        // The top‑level regions are stored in the Regions collection.
        IList<MailMergeRegionInfo> topRegions = hierarchy.Regions;

        // Iterate through each region and output its start/end field information.
        foreach (MailMergeRegionInfo region in topRegions)
        {
            // The start and end fields of the region.
            FieldMergeField startField = region.StartField;
            FieldMergeField endField = region.EndField;

            // Output the region name.
            Console.WriteLine($"Region Name: {region.Name}");

            // Output the field names that mark the start and end of the region.
            Console.WriteLine($"  Start Field: {startField.FieldName}");
            Console.WriteLine($"  End Field:   {endField.FieldName}");

            // For validation we can also output the position of the start and end fields
            // as the index of their first node within the document's node collection.
            NodeCollection allNodes = doc.GetChildNodes(NodeType.Any, true);
            int startIndex = allNodes.IndexOf(startField.Start);
            int endIndex = allNodes.IndexOf(endField.Start);
            Console.WriteLine($"  Start Position (node index): {startIndex}");
            Console.WriteLine($"  End Position   (node index): {endIndex}");
        }
    }
}
