using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class RowData
{
    public string Name { get; set; }
    public int Quantity { get; set; }
    public bool Highlight { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Load the DOTM template that contains a table with a conditional row block.
        Document template = new Document("Template.dotm");

        // Prepare the data source – a list of rows that will be iterated in the template.
        var rows = new List<RowData>
        {
            new RowData { Name = "Apple",  Quantity = 10, Highlight = true  },
            new RowData { Name = "Banana", Quantity = 5,  Highlight = false },
            new RowData { Name = "Cherry", Quantity = 12, Highlight = true  }
        };

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report. The template should contain a block such as:
        // <<foreach [Rows]>>
        //   <<if [Highlight]>>{shading:Yellow}<<endif>>
        //   <<[Name]>>    <<[Quantity]>>
        // <<endforeach>>
        // The name "Rows" is used to reference the data source inside the template.
        engine.BuildReport(template, rows, "Rows");

        // Save the populated document.
        template.Save("Result.docx");
    }
}
