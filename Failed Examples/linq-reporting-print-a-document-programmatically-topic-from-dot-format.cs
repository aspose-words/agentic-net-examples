// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Data;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Load the DOT template file. The constructor automatically detects the format.
        Document template = new Document("Template.dot");

        // Prepare a simple data source using LINQ (array of anonymous objects).
        var data = new[]
        {
            new { Name = "John Doe", Address = "123 Main St", City = "New York" },
            new { Name = "Jane Smith", Address = "456 Oak Ave", City = "Los Angeles" }
        };

        // Use the first record for a single‑document mail merge.
        var first = data[0];
        template.MailMerge.Execute(
            new[] { "Name", "Address", "City" },
            new object[] { first.Name, first.Address, first.City });

        // Save the merged document (using the provided Save method).
        template.Save("MergedDocument.docx");

        // Print the merged document to the default printer (using the provided Print method).
        template.Print();

        // Optional: print with more control using AsposeWordsPrintDocument.
        // var printDoc = new AsposeWordsPrintDocument(template);
        // var printerSettings = new PrinterSettings { PrintRange = PrintRange.AllPages };
        // printDoc.PrinterSettings = printerSettings;
        // printDoc.Print();
    }
}
