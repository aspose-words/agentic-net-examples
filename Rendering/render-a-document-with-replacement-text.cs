using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with a placeholder token.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, _Name_! Welcome to _Company_.");

        // Replace the placeholder tokens with actual values.
        // Simple string replace (case‑insensitive, whole word not required).
        doc.Range.Replace("_Name_", "John Doe");
        doc.Range.Replace("_Company_", "Aspose");

        // Save the resulting document.
        doc.Save("ReplacedDocument.docx");
    }
}
