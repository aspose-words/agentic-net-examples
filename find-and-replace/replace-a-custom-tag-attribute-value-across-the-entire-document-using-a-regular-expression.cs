using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words for font/color types

namespace ReplaceCustomTagAttribute
{
    public class Program
    {
        public static void Main()
        {
            // Paths for the sample input and output documents.
            const string inputPath = "sample_input.docx";
            const string outputPath = "sample_output.docx";

            // -----------------------------------------------------------------
            // 1. Create a sample document containing a custom XML-like tag.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert some paragraphs with the custom tag.
            builder.Writeln(@"Here is a custom tag: <custom attr=""OldValue1"">Content A</custom>");
            builder.Writeln(@"Another line with <custom attr=""OldValue2"">Content B</custom>.");
            builder.Writeln(@"A line without the tag should stay unchanged.");

            // Save the document so we can demonstrate loading it later.
            doc.Save(inputPath);

            // -----------------------------------------------------------------
            // 2. Load the document from the file system.
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(inputPath);

            // -----------------------------------------------------------------
            // 3. Define a regular expression that matches the attribute value.
            //    Pattern captures the part before the value and the closing quote,
            //    allowing us to replace only the value while preserving the rest.
            // -----------------------------------------------------------------
            // Example tag: <custom attr="OldValue1">
            // Regex groups:
            //   1 => '<custom' plus any whitespace and other attributes up to 'attr="'
            //   2 => the closing quote after the value
            Regex attributeRegex = new Regex(@"(<custom\s+[^>]*attr="")[^""]*(""[^>]*>)", RegexOptions.IgnoreCase);

            // Replacement string inserts the new attribute value ("NewValue") between the captured groups.
            const string newAttributeValue = "NewValue";
            string replacement = $"$1{newAttributeValue}$2";

            // -----------------------------------------------------------------
            // 4. Perform the replacement across the whole document.
            // -----------------------------------------------------------------
            FindReplaceOptions options = new FindReplaceOptions(); // default options
            int replacedCount = loadedDoc.Range.Replace(attributeRegex, replacement, options);

            // Validate that at least one replacement occurred.
            if (replacedCount == 0)
                throw new InvalidOperationException("No attribute values were replaced. Check the regex pattern.");

            // -----------------------------------------------------------------
            // 5. Save the modified document.
            // -----------------------------------------------------------------
            loadedDoc.Save(outputPath);

            // Optional: Write a short confirmation to the console (no user interaction required).
            Console.WriteLine($"Replaced {replacedCount} attribute value(s). Output saved to '{outputPath}'.");
        }
    }
}
