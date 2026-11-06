using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple table with a header row.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Fruit");
        builder.EndRow();
        builder.EndTable();

        // Create a repeating section content control (block level).
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "FruitRepeatingSection",
            Tag = "fruit-section"
        };

        // Inside the repeating section, create a table with a single placeholder row.
        Table innerTable = new Table(doc);
        Row placeholderRow = new Row(doc);
        Cell cell = new Cell(doc);
        Paragraph para = new Paragraph(doc);
        Run placeholderRun = new Run(doc, "<<Item>>");
        para.AppendChild(placeholderRun);
        cell.AppendChild(para);
        placeholderRow.AppendChild(cell);
        innerTable.AppendChild(placeholderRow);
        repeatingSection.AppendChild(innerTable);

        // Append the repeating section to the document body after the header table.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Sample collection to repeat the row for.
        List<string> fruits = new List<string> { "Apple", "Banana", "Cherry" };

        // Locate the repeating section and its inner table.
        StructuredDocumentTag foundRepeating = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(sdt => sdt.SdtType == SdtType.RepeatingSection);

        if (foundRepeating == null)
        {
            throw new InvalidOperationException("Repeating section not found.");
        }

        Table tableInSection = (Table)foundRepeating.GetChildNodes(NodeType.Table, true).FirstOrDefault();
        if (tableInSection == null)
        {
            throw new InvalidOperationException("Inner table not found in repeating section.");
        }

        // The first row is the placeholder row.
        Row templateRow = tableInSection.Rows[0];
        // Clear any existing rows (including the placeholder) before adding clones.
        tableInSection.Rows.Clear();

        // Clone the template row for each item and replace the placeholder text.
        foreach (string fruit in fruits)
        {
            Row newRow = (Row)templateRow.Clone(true);
            foreach (Run run in newRow.GetChildNodes(NodeType.Run, true).OfType<Run>())
            {
                if (run.Text.Contains("<<Item>>"))
                {
                    run.Text = fruit;
                }
            }
            tableInSection.Rows.Add(newRow);
        }

        // Save the resulting document.
        string docPath = "RepeatingSectionTable.docx";
        doc.Save(docPath);

        // Serialize the collection to JSON for verification.
        string jsonPath = "fruits.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(fruits, Formatting.Indented));
    }
}
