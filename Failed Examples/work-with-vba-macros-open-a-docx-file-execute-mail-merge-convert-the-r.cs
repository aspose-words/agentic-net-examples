// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Data;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        const string sourceDocPath = "Template.docx";

        // Load the existing document.
        Document doc = new Document(sourceDocPath);

        // Prepare mail‑merge data in a DataTable.
        DataTable table = new DataTable("MergeData");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Columns.Add("Address");
        table.Rows.Add("John", "Doe", "123 Main St.");
        table.Rows.Add("Jane", "Smith", "456 Oak Ave.");

        // Execute mail merge using the DataTable.
        doc.MailMerge.Execute(table);

        // Save the merged result as PDF.
        const string pdfPath = "MergedResult.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Example LINQ query: find all paragraphs that contain the word "John".
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                            .Cast<Paragraph>()
                            .Where(p => p.GetText().Contains("John"))
                            .ToList();

        Console.WriteLine($"Paragraphs containing 'John': {paragraphs.Count}");

        // Print the document to the default printer.
        doc.Print();
    }
}
