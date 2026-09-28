using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a template document in memory.
        Document template;
        using (MemoryStream templateStream = new MemoryStream())
        {
            DocumentBuilder builder = new DocumentBuilder();
            // Insert a PAGE_BREAK field before the region to start a new page.
            builder.InsertField("PAGE_BREAK", null);
            builder.Writeln();

            // Define the start of the mail merge region.
            builder.InsertField("MERGEFIELD TableStart:Employee", null);
            builder.Writeln();

            // Insert merge fields that will be populated.
            builder.InsertField("MERGEFIELD Name", null);
            builder.Writeln();
            builder.InsertField("MERGEFIELD Age", null);
            builder.Writeln();

            // Define the end of the mail merge region.
            builder.InsertField("MERGEFIELD TableEnd:Employee", null);
            builder.Writeln();

            // Save the template to a stream.
            builder.Document.Save(templateStream, SaveFormat.Docx);
            templateStream.Position = 0;
            template = new Document(templateStream);
        }

        // Prepare data for the mail merge region.
        DataTable employeeTable = new DataTable("Employee");
        employeeTable.Columns.Add("Name", typeof(string));
        employeeTable.Columns.Add("Age", typeof(int));

        employeeTable.Rows.Add("Alice", 30);
        employeeTable.Rows.Add("Bob", 25);
        employeeTable.Rows.Add("Charlie", 35);

        // Execute mail merge with regions.
        template.MailMerge.ExecuteWithRegions(employeeTable);

        // Save the result document.
        string outputPath = "Output.docx";
        template.Save(outputPath, SaveFormat.Docx);
    }
}
