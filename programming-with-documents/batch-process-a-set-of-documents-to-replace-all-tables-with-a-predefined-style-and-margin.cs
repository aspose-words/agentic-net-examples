using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Base directories for input and output documents.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create sample source documents (each contains a simple table).
        // -----------------------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            // Add a heading.
            builder.Writeln($"Sample Document {i}");

            // Build a 2x2 table with placeholder text.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();
            builder.EndTable();

            // Save the sample document.
            string samplePath = Path.Combine(inputDir, $"Sample{i}.docx");
            sampleDoc.Save(samplePath);
        }

        // -----------------------------------------------------------------
        // 2. Define a predefined table style that will be applied to all tables.
        // -----------------------------------------------------------------
        // Create a temporary document solely to hold the style definition.
        Document styleHolder = new Document();
        TableStyle predefinedStyle = (TableStyle)styleHolder.Styles.Add(StyleType.Table, "MyPredefinedStyle");
        predefinedStyle.Borders.Color = Color.Blue;
        predefinedStyle.Borders.LineStyle = LineStyle.Single;
        predefinedStyle.Borders.LineWidth = 1.5;
        predefinedStyle.CellSpacing = 5;
        predefinedStyle.BottomPadding = 10;
        predefinedStyle.TopPadding = 10;
        predefinedStyle.LeftPadding = 10;
        predefinedStyle.RightPadding = 10;
        predefinedStyle.Shading.BackgroundPatternColor = Color.LightYellow;

        // -----------------------------------------------------------------
        // 3. Process each document: apply page margins and the predefined table style.
        // -----------------------------------------------------------------
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Apply a predefined page margin setting (e.g., Narrow).
            if (doc.Sections.Count > 0)
                doc.Sections[0].PageSetup.Margins = Margins.Narrow;

            // Import the custom style from the holder document.
            doc.CopyStylesFromTemplate(styleHolder);

            // Retrieve the imported style from the current document.
            Style importedStyle = doc.Styles["MyPredefinedStyle"];

            // Iterate over all tables and assign the predefined style.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table tbl in tables)
            {
                tbl.Style = importedStyle; // Now the style belongs to this document.
                tbl.LeftIndent = 0;        // Reset left indent to align with page margins.
            }

            // Save the processed document to the output folder.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }

        // -----------------------------------------------------------------
        // 4. Simple verification (optional): list processed files.
        // -----------------------------------------------------------------
        Console.WriteLine("Batch processing completed. Processed files:");
        foreach (string outFile in Directory.GetFiles(outputDir, "*.docx"))
        {
            Console.WriteLine(Path.GetFileName(outFile));
        }
    }
}
