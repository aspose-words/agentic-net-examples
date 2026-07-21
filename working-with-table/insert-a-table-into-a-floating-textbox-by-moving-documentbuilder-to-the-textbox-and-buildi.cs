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
            Shape textbox = builder.InsertShape(ShapeType.TextBox, 300, 200);
            // Make the textbox floating and position it at the center of the page.
            textbox.WrapType = WrapType.None;
            textbox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textbox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textbox.HorizontalAlignment = HorizontalAlignment.Center;
            textbox.VerticalAlignment = VerticalAlignment.Center;

            // Move the builder's cursor to the first paragraph inside the textbox.
            builder.MoveTo(textbox.FirstParagraph);

            // Build a 2x2 table inside the floating textbox.
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

            // Save the document.
            const string outputPath = "FloatingTextboxTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not created.");
        }
    }
}
