using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableReport
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First table with LightShading style.
            Table table1 = builder.StartTable();
            // Header row.
            builder.InsertCell();
            builder.Writeln("Product");
            builder.InsertCell();
            builder.Writeln("Quantity");
            builder.EndRow();

            // Data rows.
            builder.InsertCell();
            builder.Writeln("Apples");
            builder.InsertCell();
            builder.Writeln("30");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("Bananas");
            builder.InsertCell();
            builder.Writeln("45");
            builder.EndRow();

            builder.EndTable();

            // Apply distinct style.
            table1.StyleIdentifier = StyleIdentifier.LightShading;
            table1.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;
            table1.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Add a blank paragraph to create consistent spacing between tables.
            builder.Writeln();

            // Second table with MediumShading1Accent1 style.
            Table table2 = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Country");
            builder.InsertCell();
            builder.Writeln("Capital");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("France");
            builder.InsertCell();
            builder.Writeln("Paris");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("Japan");
            builder.InsertCell();
            builder.Writeln("Tokyo");
            builder.EndRow();

            builder.EndTable();

            table2.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;
            table2.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;
            table2.AutoFit(AutoFitBehavior.AutoFitToContents);

            builder.Writeln();

            // Third table with TableGrid style.
            Table table3 = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Year");
            builder.InsertCell();
            builder.Writeln("Event");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("2020");
            builder.InsertCell();
            builder.Writeln("Olympics (Cancelled)");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("2021");
            builder.InsertCell();
            builder.Writeln("Tokyo Olympics");
            builder.EndRow();

            builder.EndTable();

            table3.StyleIdentifier = StyleIdentifier.TableGrid;
            table3.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;
            table3.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportWithMultipleTables.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The report file was not created.");

            // Optional: inform that the process completed.
            Console.WriteLine("Report generated successfully at: " + outputPath);
        }
    }
}
