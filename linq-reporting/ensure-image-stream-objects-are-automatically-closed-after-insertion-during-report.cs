using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Create a simple template document with merge fields for FirstName and an image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Employee:");
        builder.InsertField("MERGEFIELD FirstName \\* MERGEFORMAT");
        builder.Writeln();
        builder.InsertField("MERGEFIELD Photo \\* MERGEFORMAT");

        // Prepare a data source. The Photo column holds raw image bytes.
        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("FirstName", typeof(string));
        employees.Columns.Add("Photo", typeof(byte[]));

        // Tiny 1x1 pixel PNG (transparent) encoded as base64.
        const string pngBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(pngBase64);

        // Add sample rows.
        employees.Rows.Add("John", pngBytes);
        employees.Rows.Add("Jane", pngBytes);

        // Perform mail merge. Aspose.Words will automatically handle the image byte arrays.
        doc.MailMerge.Execute(employees);

        // Save the result.
        doc.Save("Report.docx");
    }
}
