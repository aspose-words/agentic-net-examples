// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains MERGEFIELDs.
        Document template = new Document("Template.docx");

        // Optional: configure mail‑merge settings to generate a new document.
        MailMergeSettings settings = template.MailMergeSettings;
        settings.Destination = MailMergeDestination.NewDocument;

        // Define the merge fields and their corresponding values.
        string[] fieldNames = { "FirstName", "LastName", "Message" };
        object[] fieldValues = { "John", "Doe", "Hello from Aspose.Words!" };

        // Execute the mail merge operation.
        template.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document as PDF.
        string pdfPath = "MergedResult.pdf";
        template.Save(pdfPath, SaveFormat.Pdf);

        // Render a simple LINQ report: count paragraphs that contain the word "Hello".
        var helloParagraphs = template.GetChildNodes(NodeType.Paragraph, true)
                                      .Cast<Paragraph>()
                                      .Where(p => p.ToTxt().Contains("Hello"));
        int count = helloParagraphs.Count();
        Console.WriteLine($"Paragraphs containing 'Hello': {count}");

        // Print the merged document to the default printer.
        template.Print();
    }
}
