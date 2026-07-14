using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace RangeFormFieldValidation
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add sample text that will later be replaced.
            builder.Writeln("Dear _Name_,");
            builder.Writeln("Welcome to the company.");

            // Insert a checkbox form field – this makes the range contain a form field.
            builder.InsertCheckBox("AcceptTerms", false, 0);

            // Use the full Aspose.Words.Range type to avoid conflict with System.Range.
            Aspose.Words.Range range = doc.Range;

            // Validate that the range does NOT contain any form fields before replacement.
            if (range.FormFields.Count == 0)
            {
                // No form fields – safe to replace.
                int replacements = range.Replace("_Name_", "John Doe");
                Console.WriteLine($"Replacement performed. Count: {replacements}");
            }
            else
            {
                // Form fields are present – skip replacement.
                Console.WriteLine($"Range contains {range.FormFields.Count} form field(s). Replacement skipped.");
            }

            // Save the resulting document.
            const string outputPath = "Result.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to '{outputPath}'.");
        }
    }
}
