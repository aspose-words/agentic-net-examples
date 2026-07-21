using System;
using System.IO;
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

            // Perform the combined insertion.
            InsertTableParagraphAndLinkedTextBox(doc);

            // Save the document to the output file.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
            doc.Save(outputPath);
        }

        /// <summary>
        /// Inserts a table, a paragraph, and a linked text box into the provided document.
        /// </summary>
        /// <param name="doc">The document to modify.</param>
        private static void InsertTableParagraphAndLinkedTextBox(Document doc)
        {
            // Use DocumentBuilder for convenient insertion.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------- Insert a simple 2x2 table ----------
            builder.StartTable();

            // First row
            builder.InsertCell();
            builder.Write("Cell 1,1");
            builder.InsertCell();
            builder.Write("Cell 1,2");
            builder.EndRow();

            // Second row
            builder.InsertCell();
            builder.Write("Cell 2,1");
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndTable();

            // Add a paragraph after the table.
            builder.Writeln();
            builder.Writeln("This paragraph follows the table and demonstrates normal text insertion.");

            // ---------- Insert a linked (floating) text box ----------
            // Create a floating text box shape.
            Shape textBox = new Shape(doc, ShapeType.TextBox);
            textBox.WrapType = WrapType.None; // Floating, not inline.
            textBox.Width = 200;
            textBox.Height = 100;
            textBox.HorizontalAlignment = HorizontalAlignment.Center;
            textBox.VerticalAlignment = VerticalAlignment.Top;

            // Add a paragraph and run inside the text box.
            Paragraph tbParagraph = new Paragraph(doc);
            Run tbRun = new Run(doc, "Content of the linked text box.");
            tbParagraph.AppendChild(tbRun);
            textBox.AppendChild(tbParagraph);

            // Insert the text box into the document at the current builder position.
            builder.InsertNode(textBox);
        }
    }
}
