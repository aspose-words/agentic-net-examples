using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsFieldSumExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a DocumentBuilder for editing.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table with two columns.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Number");
            builder.InsertCell();
            builder.Write("Description");
            builder.EndRow();

            // Data rows. Place a bookmark around the first column cells to sum them.
            for (int i = 1; i <= 3; i++)
            {
                builder.InsertCell();

                // Start bookmark on the first cell of the column.
                if (i == 1)
                    builder.StartBookmark("ColSum");

                builder.Write((i * 10).ToString());

                // End bookmark on the last cell of the column.
                if (i == 3)
                    builder.EndBookmark("ColSum");

                builder.InsertCell();
                builder.Write($"Item {i}");
                builder.EndRow();
            }

            builder.EndTable();

            // Insert a paragraph after the table.
            builder.Writeln();
            builder.Writeln("Sum of the \"Number\" column:");

            // Insert a formula field that calculates the sum of the bookmarked cells.
            // Use the string overload of InsertField to specify the field code directly.
            builder.InsertField("=SUM(ColSum)");

            // Update fields to calculate the result.
            doc.UpdateFields();

            // Save the document.
            const string outputPath = "Output.docx";
            doc.Save(outputPath);
        }
    }
}
