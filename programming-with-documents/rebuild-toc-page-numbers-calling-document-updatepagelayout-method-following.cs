using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Fields;

namespace TocPageNumberRebuilder
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "Input.docx";
            const string outputPath = "Output.docx";

            // Ensure an input file exists – create a minimal one if it does not.
            if (!File.Exists(inputPath))
            {
                var tempDoc = new Document();
                var builder = new DocumentBuilder(tempDoc);
                builder.Writeln("This is a placeholder document.");
                // Insert a simple TOC so the example has something to work with.
                builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
                tempDoc.Save(inputPath);
            }

            // Load the document.
            Document doc = new Document(inputPath);

            // Update fields that are not dependent on page layout.
            doc.UpdateFields();

            // Rebuild the page layout. Wrap in try/catch to avoid crashes caused by
            // unsupported field types (e.g., barcode fields without a license).
            try
            {
                doc.UpdatePageLayout();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: UpdatePageLayout failed – {ex.Message}");
            }

            // Update TOC page numbers.
            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldTOC)
                {
                    if (field is FieldToc toc)
                    {
                        toc.UpdatePageNumbers();
                    }
                }
            }

            // Save the updated document.
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to '{outputPath}'.");
        }
    }
}
