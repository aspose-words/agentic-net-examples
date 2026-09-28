using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListIndentExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a numbered list.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");

            // Increase indent to create a sub‑list.
            builder.ListFormat.ListLevelNumber++; // equivalent to IncreaseIndent
            builder.Writeln("Subitem 2.1");

            // Conditionally decrease indent if we are deeper than the top level.
            if (builder.ListFormat.ListLevelNumber > 0)
            {
                builder.ListFormat.ListLevelNumber--; // equivalent to DecreaseIndent
            }

            // Continue at the original list level.
            builder.Writeln("Item 3");

            // Save the document.
            doc.Save("ListIndentExample.docx");
        }
    }
}
