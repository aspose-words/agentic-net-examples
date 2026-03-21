using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMailMergeToPdf
{
    class Program
    {
        static void Main()
        {
            // Create a simple DOCX template with MERGEFIELDs if it doesn't already exist.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            CreateTemplateIfMissing(templatePath);

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare a simple DataTable with data for the mail merge.
            DataTable table = new DataTable("Employees");
            table.Columns.Add("FullName");
            table.Columns.Add("Company");
            table.Columns.Add("Address");
            table.Columns.Add("City");

            table.Rows.Add("James Bond", "MI5 Headquarters", "Milbank", "London");
            table.Rows.Add("Ethan Hunt", "IMF", "123 Secret St.", "Washington");

            // Execute the mail merge using the DataTable.
            doc.MailMerge.Execute(table);

            // Save the merged document as PDF in the current directory.
            string pdfOutputPath = Path.Combine(Environment.CurrentDirectory, "MergedResult.pdf");
            doc.Save(pdfOutputPath, SaveFormat.Pdf);

            Console.WriteLine($"Mail merge completed and PDF saved to: {pdfOutputPath}");
        }

        private static void CreateTemplateIfMissing(string path)
        {
            if (File.Exists(path))
                return;

            Document template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Employee Details:");
            builder.Writeln();

            builder.InsertField("MERGEFIELD FullName \\* MERGEFORMAT");
            builder.Writeln();

            builder.InsertField("MERGEFIELD Company \\* MERGEFORMAT");
            builder.Writeln();

            builder.InsertField("MERGEFIELD Address \\* MERGEFORMAT");
            builder.Writeln();

            builder.InsertField("MERGEFIELD City \\* MERGEFORMAT");
            builder.Writeln();

            template.Save(path);
        }
    }
}
