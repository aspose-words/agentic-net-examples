using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;   // Needed for the Table class

namespace FloatingTextboxTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a floating textbox shape.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
            // Optional: set floating positioning.
            textBox.WrapType = WrapType.None;
            textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBox.HorizontalAlignment = HorizontalAlignment.Right;
            textBox.VerticalAlignment = VerticalAlignment.Bottom;
            textBox.Left = 50;
            textBox.Top = 50;

            // Move the builder cursor into the textbox's first paragraph.
            builder.MoveTo(textBox.FirstParagraph);

            // Build a 2x2 table inside the textbox.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FloatingTextboxTable.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
