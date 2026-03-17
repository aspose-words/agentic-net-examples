using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("input.docx");

        // Obtain a specific range – here we use the range of the first section.
        // Use 'var' (or an alias) to avoid the ambiguous 'Range' type introduced in .NET 5+.
        var range = doc.FirstSection.Range;

        // Retrieve the number of form fields that exist within this range.
        int formFieldCount = range.FormFields.Count;

        // Output the count.
        Console.WriteLine($"Form fields in the selected range: {formFieldCount}");

        // Save the document (optional, demonstrates the required save step).
        doc.Save("output.docx");
    }
}
