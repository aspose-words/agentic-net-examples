using System;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace RangeFormFieldValidation
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add sample text that we intend to replace later.
            builder.Writeln("Hello _Name_!");

            // Validate that the whole document range contains no form fields.
            // The FormFields collection is available on the Range object.
            if (doc.Range.FormFields.Count == 0)
            {
                // Since there are no form fields, perform the replacement.
                int replacements = doc.Range.Replace("_Name_", "World");
                Console.WriteLine($"Replacements made: {replacements}");
            }
            else
            {
                Console.WriteLine("The range contains form fields; replacement skipped.");
            }

            // Save the resulting document.
            doc.Save("Result.docx");

            // Output the final document text to the console for verification.
            Console.WriteLine("Final document text:");
            Console.WriteLine(doc.GetText().Trim());
        }
    }
}
