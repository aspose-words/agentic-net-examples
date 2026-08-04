using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace AsposeWordsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Attach a DocumentBuilder to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the required elements in one operation.
            InsertElements(builder);

            // Save the document to a file.
            doc.Save("MacroLikeOutput.docx");
        }

        /// <summary>
        /// Inserts a 2x2 table, a paragraph, and a linked text box into the document.
        /// </summary>
        private static void InsertElements(DocumentBuilder builder)
        {
            // ----- Insert a 2x2 table -----
            builder.StartTable();

            // First row
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Second row
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndTable();

            // Add a paragraph after the table.
            builder.Writeln("This is a paragraph following the table.");

            // ----- Insert a linked text box -----
            // Create a floating text box shape.
            Shape textBox = new Shape(builder.Document, ShapeType.TextBox);
            textBox.WrapType = WrapType.None;
            textBox.Width = 200;
            textBox.Height = 100;

            // Add a paragraph with some text inside the text box.
            Paragraph tbParagraph = new Paragraph(builder.Document);
            Run tbRun = new Run(builder.Document, "Content of the linked text box.");
            tbParagraph.AppendChild(tbRun);
            textBox.AppendChild(tbParagraph);

            // Insert the text box into the document.
            builder.InsertNode(textBox);

            // Add a paragraph after the linked text box.
            builder.Writeln("Paragraph after the linked text box.");
        }
    }
}
