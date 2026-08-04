using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a mail merge region named "MyRegion" with two fields.
        builder.InsertField(" MERGEFIELD TableStart:MyRegion");
        builder.InsertField(" MERGEFIELD Field1");
        builder.Write(", ");
        builder.InsertField(" MERGEFIELD Field2");
        builder.InsertField(" MERGEFIELD TableEnd:MyRegion");

        // Retrieve the hierarchy of mail merge regions.
        MailMergeRegionInfo hierarchy = doc.MailMerge.GetRegionsHierarchy();

        // Get the top‑level regions from the hierarchy.
        IList<MailMergeRegionInfo> topRegions = hierarchy.Regions;

        if (topRegions.Count > 0)
        {
            MailMergeRegionInfo region = topRegions[0];

            // Access the start and end fields of the region.
            FieldMergeField startField = region.StartField;
            FieldMergeField endField = region.EndField;

            Console.WriteLine("Region name: " + region.Name);
            Console.WriteLine("Start field name: " + startField.FieldName);
            Console.WriteLine("End field name: " + endField.FieldName);
        }
        else
        {
            Console.WriteLine("No mail merge regions found.");
        }

        // Save the document (optional, demonstrates the save lifecycle).
        doc.Save("MailMergeRegionInfo.docx");
    }
}
