using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields; // Added for Field type

namespace MailMergeRegionInfoDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source document that contains mail merge regions.
            // (Assumes the file exists at the specified location.)
            Document doc = new Document("MailMergeRegions.docx");

            // Retrieve the full hierarchy of mail merge regions.
            MailMergeRegionInfo hierarchy = doc.MailMerge.GetRegionsHierarchy();

            // Process each top‑level region.
            IList<MailMergeRegionInfo> topRegions = hierarchy.Regions;
            foreach (MailMergeRegionInfo topRegion in topRegions)
            {
                PrintRegionInfo(topRegion, indentLevel: 0);
            }

            // Save the document if any modifications were made (not required for this demo).
            // doc.Save("MailMergeRegions_Output.docx");
        }

        // Recursively prints region metadata, including start/end field names.
        private static void PrintRegionInfo(MailMergeRegionInfo region, int indentLevel)
        {
            string indent = new string(' ', indentLevel * 4);

            // Region name and nesting level.
            Console.WriteLine($"{indent}Region: {region.Name}, Level: {region.Level}");

            // Start and end merge fields that delimit the region.
            FieldMergeField startField = region.StartField;
            FieldMergeField endField = region.EndField;

            // Verify that start and end fields are present.
            if (startField != null && endField != null)
            {
                Console.WriteLine($"{indent}    Start Field: {startField.FieldName}");
                Console.WriteLine($"{indent}    End   Field: {endField.FieldName}");
            }
            else
            {
                Console.WriteLine($"{indent}    Warning: Missing start or end field.");
            }

            // List child fields inside the region.
            IList<Field> fields = region.Fields;
            Console.WriteLine($"{indent}    Fields count: {fields.Count}");

            // Recursively handle nested regions.
            IList<MailMergeRegionInfo> childRegions = region.Regions;
            foreach (MailMergeRegionInfo child in childRegions)
            {
                PrintRegionInfo(child, indentLevel + 1);
            }
        }
    }
}
