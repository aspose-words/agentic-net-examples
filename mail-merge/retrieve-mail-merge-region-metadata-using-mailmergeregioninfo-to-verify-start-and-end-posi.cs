using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;   // Needed for FieldMergeField

namespace MailMergeRegionInfoExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define a mail merge region named "MyRegion" with two fields.
            builder.InsertField(" MERGEFIELD TableStart:MyRegion");
            builder.InsertField(" MERGEFIELD FirstName");
            builder.Write(" ");
            builder.InsertField(" MERGEFIELD LastName");
            builder.InsertField(" MERGEFIELD TableEnd:MyRegion");

            // Retrieve the full hierarchy of mail merge regions.
            MailMergeRegionInfo hierarchy = doc.MailMerge.GetRegionsHierarchy();

            // The top‑level regions are stored in the Regions collection.
            IList<MailMergeRegionInfo> topRegions = hierarchy.Regions;

            // Output information about each region.
            foreach (MailMergeRegionInfo region in topRegions)
            {
                Console.WriteLine($"Region Name: {region.Name}");
                Console.WriteLine($"Nesting Level: {region.Level}");

                // StartField and EndField give the MERGEFIELD objects that mark the region.
                FieldMergeField startField = region.StartField;
                FieldMergeField endField = region.EndField;

                Console.WriteLine($"Start Field Name: {startField?.FieldName}");
                Console.WriteLine($"End Field Name: {endField?.FieldName}");
                Console.WriteLine(new string('-', 40));
            }

            // Save the document to the current directory (optional, just to demonstrate saving).
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MailMergeRegionInfoOutput.docx");
            doc.Save(outputPath);
        }
    }
}
