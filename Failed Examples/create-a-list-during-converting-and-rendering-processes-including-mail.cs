// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Data;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("input.docx");

        // Generate mail merge data using LINQ.
        var customers = new[]
        {
            new { Name = "Alice", Address = "123 Main St" },
            new { Name = "Bob",   Address = "456 Oak Ave" }
        };

        DataTable table = new DataTable("Customers");
        table.Columns.Add("Name");
        table.Columns.Add("Address");
        foreach (var c in customers)
        {
            table.Rows.Add(c.Name, c.Address);
        }

        // Perform mail merge.
        doc.MailMerge.Execute(table);

        // Create a list to capture process information.
        List<string> log = new List<string>();
        log.Add($"Mail merge executed with {table.Rows.Count} rows.");

        // Render each page of the document to an image file.
        int pageCount = doc.PageCount;
        for (int i = 0; i < pageCount; i++)
        {
            using (Bitmap bmp = new Bitmap(800, 1000))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    // Render the current page at 100% scale.
                    doc.RenderToScale(i, g, 1.0f);
                }

                string imagePath = $"page_{i + 1}.png";
                bmp.Save(imagePath);
                log.Add($"Rendered page {i + 1} to {imagePath}");
            }
        }

        // Print the document.
        doc.Print();

        // Save the final document after mail merge.
        doc.Save("output.docx");
        log.Add("Document saved as output.docx");

        // Output the log to the console.
        foreach (string entry in log)
        {
            Console.WriteLine(entry);
        }
    }
}
