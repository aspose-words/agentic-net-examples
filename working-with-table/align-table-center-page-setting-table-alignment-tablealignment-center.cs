using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class AlignTableCenter
{
    static void Main()
    {
        const string inputPath = "Input.docx";
        const string outputPath = "Output.docx";

        Document doc;

        // If the input file exists, load it; otherwise, create a new document with a sample table.
        if (File.Exists(inputPath))
        {
            doc = new Document(inputPath);
        }
        else
        {
            doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Sample document created because Input.docx was not found.");
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();
            builder.EndTable();
        }

        // Ensure the document has at least one section and one table.
        if (doc.FirstSection?.Body?.Tables?.Count > 0)
        {
            Table table = doc.FirstSection.Body.Tables[0];
            table.Alignment = TableAlignment.Center;
        }
        else
        {
            Console.WriteLine("The document does not contain any tables to align.");
        }

        // Save the modified document.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
