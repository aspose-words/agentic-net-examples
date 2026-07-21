using System;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a mail‑merge region named "MyRegion" with two data fields.
        builder.InsertField(" MERGEFIELD TableStart:MyRegion");
        builder.InsertField(" MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField(" MERGEFIELD LastName");
        builder.InsertField(" MERGEFIELD TableEnd:MyRegion");

        // Retrieve the full hierarchy of mail‑merge regions in the document.
        MailMergeRegionInfo hierarchy = doc.MailMerge.GetRegionsHierarchy();

        // Iterate over top‑level regions (there is only one in this example).
        foreach (MailMergeRegionInfo region in hierarchy.Regions)
        {
            Console.WriteLine($"Region name: {region.Name}");
            Console.WriteLine($"Nesting level: {region.Level}");

            // The StartField and EndField mark the boundaries of the region.
            FieldMergeField startField = region.StartField;
            FieldMergeField endField = region.EndField;

            Console.WriteLine($"Start field name: {startField?.FieldName}");
            Console.WriteLine($"End field name: {endField?.FieldName}");

            // Simple validation: the field names must contain the region name with the correct prefix.
            bool startValid = startField?.FieldName == $"TableStart:{region.Name}";
            bool endValid = endField?.FieldName == $"TableEnd:{region.Name}";

            Console.WriteLine($"Start tag valid: {startValid}");
            Console.WriteLine($"End tag valid: {endValid}");
        }

        // Save the document to disk (optional, demonstrates that the document is valid).
        string outputPath = "MailMergeRegionInfoExample.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
