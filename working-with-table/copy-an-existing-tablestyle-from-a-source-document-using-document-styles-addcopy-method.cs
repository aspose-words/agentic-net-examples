using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableStyleCopy
{
    public class Program
    {
        public static void Main()
        {
            // Define output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -------------------------
            // Create source document with a custom table style.
            // -------------------------
            Document sourceDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

            // Create a custom table style.
            TableStyle sourceStyle = (TableStyle)sourceDoc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
            sourceStyle.Shading.BackgroundPatternColor = Color.AntiqueWhite;
            sourceStyle.Borders.Color = Color.Blue;
            sourceStyle.Borders.LineStyle = LineStyle.Single;
            sourceStyle.CellSpacing = 5;
            sourceStyle.VerticalAlignment = CellVerticalAlignment.Center;

            // Build a simple table and apply the custom style.
            Table srcTable = srcBuilder.StartTable();
            srcBuilder.InsertCell();
            srcBuilder.Write("Source Cell 1");
            srcBuilder.InsertCell();
            srcBuilder.Write("Source Cell 2");
            srcBuilder.EndRow();
            srcBuilder.EndTable();

            srcTable.Style = sourceStyle;

            // Save the source document.
            string sourcePath = Path.Combine(outputDir, "Source.docx");
            sourceDoc.Save(sourcePath);

            // -------------------------
            // Create destination document and copy the table style from the source.
            // -------------------------
            Document destDoc = new Document();

            // Copy the style using the AddCopy method.
            Style copiedStyle = destDoc.Styles.AddCopy(sourceStyle);
            // Optionally rename the copied style.
            copiedStyle.Name = "CopiedTableStyle";

            // Build a table in the destination document and apply the copied style.
            DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
            Table destTable = destBuilder.StartTable();
            destBuilder.InsertCell();
            destBuilder.Write("Dest Cell 1");
            destBuilder.InsertCell();
            destBuilder.Write("Dest Cell 2");
            destBuilder.EndRow();
            destBuilder.EndTable();

            destTable.Style = copiedStyle;

            // Save the destination document.
            string destPath = Path.Combine(outputDir, "Destination.docx");
            destDoc.Save(destPath);

            // Simple verification (throws if the style was not copied correctly).
            if (destTable.StyleName != copiedStyle.Name)
                throw new InvalidOperationException("The table style was not applied correctly.");

            // Indicate completion.
            Console.WriteLine("Source and destination documents have been created successfully.");
        }
    }
}
