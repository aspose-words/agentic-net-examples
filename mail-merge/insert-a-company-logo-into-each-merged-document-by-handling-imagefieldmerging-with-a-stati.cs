using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a temporary 1x1 pixel PNG logo.
        string logoPath = Path.Combine(Path.GetTempPath(), "logo.png");
        byte[] pngData = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X3eUAAAAASUVORK5CYII=");
        File.WriteAllBytes(logoPath, pngData);

        // Build a simple template with a text merge field and an image merge field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear <<Name>>,");
        builder.InsertField("MERGEFIELD CompanyLogo \\* MERGEFORMAT");
        builder.Writeln();
        builder.Writeln("Thank you for your business.");

        // Prepare a data source containing a Name field.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Alice");
        table.Rows.Add("Bob");

        // Execute mail merge.
        doc.MailMerge.Execute(table);

        // After merging, replace each CompanyLogo merge field with the static logo image.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldMergeField && field.GetFieldCode().Contains("CompanyLogo"))
            {
                DocumentBuilder imgBuilder = new DocumentBuilder(doc);
                imgBuilder.MoveTo(field.Start);
                imgBuilder.InsertImage(logoPath);
                field.Remove(); // Remove the original merge field.
            }
        }

        // Save the merged document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocument.docx");
        doc.Save(outputPath);

        // Clean up the temporary logo file.
        File.Delete(logoPath);
    }
}
