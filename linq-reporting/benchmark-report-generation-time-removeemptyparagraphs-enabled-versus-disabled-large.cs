using System;
using System.Data;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.MailMerging;

class ReportGenerationBenchmark
{
    static void Main()
    {
        // Path for the temporary template and result files.
        const string templatePath = "ReportTemplate.docx";
        const string resultWithCleanupPath = "Report_WithRemoveEmptyParagraphs.docx";
        const string resultWithoutCleanupPath = "Report_WithoutRemoveEmptyParagraphs.docx";

        // -----------------------------------------------------------------
        // 1. Create a large template document containing merge fields.
        // -----------------------------------------------------------------
        Document template = new Document();                     // create
        DocumentBuilder builder = new DocumentBuilder(template);
        // Add a large number of paragraphs with merge fields.
        for (int i = 0; i < 10000; i++)
        {
            // Some fields will be left empty during the merge to generate empty paragraphs.
            builder.Writeln($"Paragraph {i + 1}: <<FirstName>> <<LastName>>");
        }
        template.Save(templatePath);                           // save

        // -----------------------------------------------------------------
        // 2. Prepare a data source where the merge fields are empty.
        // -----------------------------------------------------------------
        DataTable data = new DataTable("Employees");
        data.Columns.Add("FirstName", typeof(string));
        data.Columns.Add("LastName", typeof(string));
        // Insert a single row with empty strings – this will produce empty paragraphs.
        data.Rows.Add(string.Empty, string.Empty);

        // -----------------------------------------------------------------
        // 3. Benchmark with RemoveEmptyParagraphs enabled.
        // -----------------------------------------------------------------
        Document docWithCleanup = new Document(templatePath);   // load
        // Enable the cleanup option that removes empty paragraphs after the merge.
        docWithCleanup.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
        Stopwatch swWith = Stopwatch.StartNew();
        docWithCleanup.MailMerge.Execute(data);
        swWith.Stop();
        docWithCleanup.Save(resultWithCleanupPath);            // save

        // -----------------------------------------------------------------
        // 4. Benchmark with RemoveEmptyParagraphs disabled.
        // -----------------------------------------------------------------
        Document docWithoutCleanup = new Document(templatePath); // load
        // Ensure the cleanup option is not set.
        docWithoutCleanup.MailMerge.CleanupOptions = MailMergeCleanupOptions.None;
        Stopwatch swWithout = Stopwatch.StartNew();
        docWithoutCleanup.MailMerge.Execute(data);
        swWithout.Stop();
        docWithoutCleanup.Save(resultWithoutCleanupPath);      // save

        // -----------------------------------------------------------------
        // 5. Output the measured times.
        // -----------------------------------------------------------------
        Console.WriteLine($"Time with RemoveEmptyParagraphs: {swWith.ElapsedMilliseconds} ms");
        Console.WriteLine($"Time without RemoveEmptyParagraphs: {swWithout.ElapsedMilliseconds} ms");
    }
}
