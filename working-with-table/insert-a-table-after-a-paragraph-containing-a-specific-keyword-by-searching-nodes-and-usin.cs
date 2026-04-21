using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableInsertExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some paragraphs – one of them contains the keyword we will search for.
            builder.Writeln("This is the first paragraph.");
            builder.Writeln("Please insert the table after this paragraph: INSERT_TABLE_HERE");
            builder.Writeln("This is the last paragraph.");

            // Search for the paragraph that contains the specific keyword.
            Paragraph keywordParagraph = null;
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                if (para.GetText().Contains("INSERT_TABLE_HERE"))
                {
                    keywordParagraph = para;
                    break;
                }
            }

            if (keywordParagraph == null)
                throw new InvalidOperationException("Keyword paragraph not found.");

            // Build a simple 2x2 table using DocumentBuilder (preferred workflow).
            // Move the builder to the paragraph after which we want to insert the table.
            builder.MoveTo(keywordParagraph);
            // Insert a new empty paragraph so the table will be placed after the keyword paragraph.
            builder.Writeln();

            // Start the table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1,1");
            builder.InsertCell();
            builder.Write("Cell 1,2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 2,1");
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the resulting document.
            string outputPath = "Result.docx";
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new IOException($"Failed to create the output file: {outputPath}");
        }
    }
}
