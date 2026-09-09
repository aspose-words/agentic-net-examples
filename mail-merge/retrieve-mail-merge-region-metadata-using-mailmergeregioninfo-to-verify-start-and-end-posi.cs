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
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a simple mail merge region named "Region1" with two fields.
            builder.InsertField(" MERGEFIELD TableStart:Region1");
            builder.InsertField(" MERGEFIELD Column1");
            builder.InsertField(" MERGEFIELD TableEnd:Region1");

            // Retrieve the full hierarchy of mail merge regions.
            MailMergeRegionInfo hierarchy = doc.MailMerge.GetRegionsHierarchy();

            // The top-level regions are stored in the Regions collection.
            IList<MailMergeRegionInfo> topRegions = hierarchy.Regions;

            // Output information about each region.
            foreach (MailMergeRegionInfo region in topRegions)
            {
                Console.WriteLine($"Region Name: {region.Name}");
                Console.WriteLine($"Nesting Level: {region.Level}");

                // Start and end fields contain the MERGEFIELD that marks the region boundaries.
                FieldMergeField startField = region.StartField;
                FieldMergeField endField = region.EndField;

                Console.WriteLine($"Start Field Name: {startField?.FieldName}");
                Console.WriteLine($"End Field Name: {endField?.FieldName}");

                // List all child fields inside the region.
                IList<Field> fields = region.Fields;
                Console.WriteLine($"Number of child fields: {fields.Count}");
                foreach (Field f in fields)
                {
                    if (f is FieldMergeField mergeField)
                        Console.WriteLine($"  Child Field: {mergeField.FieldName}");
                }

                Console.WriteLine(new string('-', 40));
            }

            // Save the document to verify that the region was created correctly.
            doc.Save("MailMergeRegionInfoOutput.docx");
        }
    }
}
