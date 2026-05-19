using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a header row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Value");
        builder.EndRow();

        // Add a regular data row.
        builder.InsertCell();
        builder.Write("Sample");
        builder.InsertCell();
        int dataValue = 75; // Example data value.
        builder.Write(dataValue.ToString());
        builder.EndRow();

        // Add a conditional row that appears only when dataValue exceeds the threshold.
        int threshold = 50;
        builder.InsertCell();
        // Insert an IF field that displays text only if the condition is true.
        string ifFieldCode = $" IF {dataValue} > {threshold} \"Conditional Row\" \"\" ";
        builder.InsertField(ifFieldCode);
        builder.InsertCell();
        builder.Write("Above Threshold");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConditionalTable.docx");
        doc.Save(outputPath);
    }
}
