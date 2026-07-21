using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder to add mail merge fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple mail merge region named "Region1" with two fields.
        builder.InsertField(" MERGEFIELD TableStart:Region1");
        builder.InsertField(" MERGEFIELD FieldA ");
        builder.Write(", ");
        builder.InsertField(" MERGEFIELD FieldB ");
        builder.InsertField(" MERGEFIELD TableEnd:Region1");

        // Retrieve the full hierarchy of mail merge regions.
        MailMergeRegionInfo hierarchy = doc.MailMerge.GetRegionsHierarchy();

        // The top-level regions are stored in the Regions collection.
        IList<MailMergeRegionInfo> topRegions = hierarchy.Regions;

        Console.WriteLine("Mail merge region metadata:");
        foreach (MailMergeRegionInfo region in topRegions)
        {
            // Output basic region information.
            Console.WriteLine($"Region Name: {region.Name}");
            Console.WriteLine($"Nesting Level: {region.Level}");

            // Start and end fields contain the MERGEFIELD names that mark the region.
            FieldMergeField startField = region.StartField;
            FieldMergeField endField = region.EndField;

            Console.WriteLine($"Start Field Name: {startField?.FieldName ?? "None"}");
            Console.WriteLine($"End Field Name: {endField?.FieldName ?? "None"}");
            Console.WriteLine();
        }

        // Optionally, save the document to verify the region layout (not required for metadata).
        doc.Save("MailMergeRegionInfo.docx");
    }
}
