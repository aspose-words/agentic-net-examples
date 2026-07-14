using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

namespace MailMergeRegionInfoExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and define a simple mail merge region.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the start tag of the region.
            builder.InsertField(" MERGEFIELD TableStart:MyRegion");
            // Insert some merge fields that belong to the region.
            builder.InsertField(" MERGEFIELD Column1");
            builder.Write(", ");
            builder.InsertField(" MERGEFIELD Column2");
            // Insert the end tag of the region.
            builder.InsertField(" MERGEFIELD TableEnd:MyRegion");

            // Retrieve the full hierarchy of mail merge regions.
            MailMergeRegionInfo hierarchy = doc.MailMerge.GetRegionsHierarchy();
            IList<MailMergeRegionInfo> topRegions = hierarchy.Regions;

            // Output information about each region.
            foreach (MailMergeRegionInfo region in topRegions)
            {
                Console.WriteLine($"Region Name: {region.Name}");
                Console.WriteLine($"Nesting Level: {region.Level}");

                // Start and end fields of the region.
                FieldMergeField startField = region.StartField;
                FieldMergeField endField = region.EndField;

                Console.WriteLine($"Start Field Name: {startField?.FieldName}");
                Console.WriteLine($"End Field Name: {endField?.FieldName}");

                // List all child fields inside the region.
                IList<Field> fields = region.Fields;
                Console.WriteLine("Fields inside the region:");
                foreach (Field field in fields)
                {
                    if (field is FieldMergeField mergeField)
                        Console.WriteLine($"  {mergeField.FieldName}");
                }

                Console.WriteLine(new string('-', 40));
            }

            // Save the document (optional, just to demonstrate saving).
            doc.Save("MailMergeRegionInfoOutput.docx");
        }
    }
}
