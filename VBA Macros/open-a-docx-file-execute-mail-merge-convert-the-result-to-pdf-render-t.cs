using System;
using System.Linq;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX template.
        Document doc = new Document("Template.docx");

        // Execute a simple mail merge.
        string[] fieldNames = { "FullName", "Address" };
        object[] fieldValues = { "John Doe", "123 Main St" };
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Render the merged document using LINQ (extract non‑empty paragraph texts).
        var paragraphTexts = doc.GetChildNodes(NodeType.Paragraph, true)
                                .Cast<Paragraph>()
                                .Select(p => p.GetText().Trim())
                                .Where(t => !string.IsNullOrEmpty(t));

        Console.WriteLine("Merged Report Content:");
        foreach (var text in paragraphTexts)
            Console.WriteLine(text);

        // Save the merged document as PDF.
        const string pdfPath = "MergedResult.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Print the PDF document using the default printer.
        var printInfo = new ProcessStartInfo
        {
            FileName = pdfPath,
            Verb = "print",
            CreateNoWindow = true,
            UseShellExecute = true
        };
        Process.Start(printInfo);
    }
}
