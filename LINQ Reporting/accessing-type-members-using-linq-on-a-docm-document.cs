using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Properties;

class Program
{
    static void Main()
    {
        // Load an existing DOCM document.
        Document doc = new Document("Sample.docm");

        // ------------------------------------------------------------
        // Example 1: Use LINQ to retrieve all custom properties that are strings.
        // ------------------------------------------------------------
        var stringProperties = doc.CustomDocumentProperties
            .Where(p => p.Type == PropertyType.String)               // Filter by property type.
            .Select(p => new { Name = p.Name, Value = p.Value.ToString() }) // Project to a simple anonymous type.
            .ToList();

        foreach (var prop in stringProperties)
        {
            Console.WriteLine($"String Property: {prop.Name} = {prop.Value}");
        }

        // ------------------------------------------------------------
        // Example 2: Use LINQ to find all hyperlink fields in the document.
        // ------------------------------------------------------------
        var hyperlinkFields = doc.Range.Fields
            .Where(f => f.Type == FieldType.FieldHyperlink)          // Keep only hyperlink fields.
            .Select(f => new { Code = f.GetFieldCode(), Result = f.Result }) // Project field code and displayed result.
            .ToList();

        foreach (var link in hyperlinkFields)
        {
            Console.WriteLine($"Hyperlink Field: Code=\"{link.Code}\", Result=\"{link.Result}\"");
        }

        // Save the (potentially modified) document.
        doc.Save("Result.docx");
    }
}
