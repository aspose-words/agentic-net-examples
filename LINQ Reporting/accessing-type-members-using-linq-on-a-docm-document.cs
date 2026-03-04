using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Properties;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCM (macro-enabled) document.
        // The Document constructor handles the lifecycle creation and loading.
        Document doc = new Document(@"C:\Docs\Sample.docm");

        // -----------------------------------------------------------------
        // Example 1: Use LINQ to query custom document properties.
        // -----------------------------------------------------------------
        // Cast the collection to IEnumerable<DocumentProperty> so LINQ can be applied.
        var stringProperties = doc.CustomDocumentProperties
                                   .Cast<DocumentProperty>()
                                   .Where(p => p.Type == PropertyType.String);

        Console.WriteLine("Custom string properties:");
        foreach (var prop in stringProperties)
        {
            Console.WriteLine($"- {prop.Name}: {prop.Value}");
        }

        // -----------------------------------------------------------------
        // Example 2: Use LINQ to find all hyperlink fields in the document.
        // -----------------------------------------------------------------
        var hyperlinkFields = doc.Range.Fields
                                   .Cast<Field>()
                                   .Where(f => f.Type == FieldType.FieldHyperlink);

        Console.WriteLine("\nHyperlink fields:");
        foreach (var field in hyperlinkFields)
        {
            // The field code contains the target URL; extract it for display.
            string fieldCode = field.GetFieldCode();
            Console.WriteLine($"- {fieldCode}");
        }

        // Save the document after any modifications (none in this example).
        // The Save method follows the lifecycle rule for persisting the document.
        doc.Save(@"C:\Docs\Sample_Processed.docm");
    }
}
