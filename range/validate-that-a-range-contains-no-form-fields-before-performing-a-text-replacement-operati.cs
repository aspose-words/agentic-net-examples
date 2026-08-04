using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsRangeValidation
{
    public class Program
    {
        public static void Main()
        {
            // Define file paths.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);
            string sourcePath = Path.Combine(artifactsDir, "Source.docx");
            string resultPath = Path.Combine(artifactsDir, "Result.docx");

            // -----------------------------------------------------------------
            // 1. Create a sample document with placeholder text.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Dear _Name_,");
            builder.Writeln("Welcome to the Aspose.Words demo.");
            // Uncomment the following line to add a form field and see the validation fail.
            // builder.InsertCheckBox("AcceptTerms", false, false, 0);

            // Save the source document.
            doc.Save(sourcePath);

            // -----------------------------------------------------------------
            // 2. Load the document from disk.
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(sourcePath);

            // -----------------------------------------------------------------
            // 3. Validate that the document's range contains no form fields.
            // -----------------------------------------------------------------
            // Use fully qualified type name to avoid conflict with System.Range.
            Aspose.Words.Range range = loadedDoc.Range;
            if (range.FormFields.Count == 0)
            {
                // No form fields found – perform the text replacement.
                int replacements = range.Replace("_Name_", "John Doe");
                Console.WriteLine($"Replacement performed. Count: {replacements}");
            }
            else
            {
                // Form fields exist – skip replacement.
                Console.WriteLine("The range contains form fields; replacement skipped.");
            }

            // -----------------------------------------------------------------
            // 4. Save the resulting document.
            // -----------------------------------------------------------------
            loadedDoc.Save(resultPath);
            Console.WriteLine($"Result document saved to: {resultPath}");
        }
    }
}
