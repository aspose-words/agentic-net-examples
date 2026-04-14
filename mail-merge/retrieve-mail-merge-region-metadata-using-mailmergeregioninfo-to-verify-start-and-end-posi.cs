using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;   // Needed for FieldMergeField and Field

namespace MailMergeRegionInfoExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a DocumentBuilder to construct a mail merge region.
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

            // Save the document (optional, just to have a physical file).
            doc.Save("MailMergeRegionInfo.docx");

            // Retrieve the full hierarchy of mail merge regions.
            MailMergeRegionInfo hierarchy = doc.MailMerge.GetRegionsHierarchy();

            // The top‑level regions are stored in the Regions collection.
            IList<MailMergeRegionInfo> topRegions = hierarchy.Regions;

            Console.WriteLine("Mail merge region metadata:");
            foreach (MailMergeRegionInfo region in topRegions)
            {
                // Region name and nesting level.
                Console.WriteLine($"Region Name: {region.Name}");
                Console.WriteLine($"Nesting Level: {region.Level}");

                // Start and end fields of the region.
                FieldMergeField startField = region.StartField;
                FieldMergeField endField = region.EndField;

                Console.WriteLine($"Start Field Name: {startField?.FieldName}");
                Console.WriteLine($"End Field Name: {endField?.FieldName}");

                // List child fields inside the region.
                IList<Field> fields = region.Fields;
                Console.WriteLine($"Number of child fields: {fields.Count}");
                foreach (Field f in fields)
                {
                    if (f is FieldMergeField mergeField)
                        Console.WriteLine($"  Child Field: {mergeField.FieldName}");
                }

                Console.WriteLine();
            }

            // The program finishes without waiting for user input.
        }
    }
}
