using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

class MailMergeCleanupExample
{
    static void Main()
    {
        // Create a simple template document in memory with MERGEFIELDs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Employee List:");
        builder.InsertParagraph();
        builder.InsertField("MERGEFIELD FirstName", "«FirstName»");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "«LastName»");
        builder.Writeln();

        // Configure mail merge to remove any paragraphs that become empty after merging.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
        doc.MailMerge.CleanupParagraphsWithPunctuationMarks = true;

        // Prepare a simple data source.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Rows.Add(new object[] { "John", "Doe" });
        table.Rows.Add(new object[] { "", "" }); // This row will produce empty paragraphs.
        table.Rows.Add(new object[] { "Jane", "Doe" });

        // Execute the mail merge.
        doc.MailMerge.ExecuteWithRegions(table);

        // Save the cleaned‑up document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MergedResult.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Merged document saved to: {outputPath}");
    }
}
