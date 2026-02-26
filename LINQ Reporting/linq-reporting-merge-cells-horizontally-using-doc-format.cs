using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Reporting;

namespace AsposeWordsMergeExample
{
    // Simple data source class for the reporting engine.
    public class ReportData
    {
        public string Name { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a blank Word document.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // 2. Build a table where the first two cells are merged horizontally.
            //    The merged cell will contain a reporting placeholder that will be replaced
            //    by the ReportingEngine.
            builder.StartTable();

            // First cell – start of the merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            // Reporting placeholder – the engine will replace this with the value of data.Name.
            builder.Write("<<[data.Name]>>");

            // Second cell – continues the merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            // No content needed; this cell is merged with the previous one.
            builder.Write(string.Empty);

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Reset the merge flag so that any subsequent cells are not affected.
            builder.CellFormat.HorizontalMerge = CellMerge.None;

            // 3. Prepare the data source.
            var data = new ReportData { Name = "Merged Cell Content" };

            // 4. Execute the LINQ reporting engine to fill the placeholder.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, data, "data");

            // 5. Save the resulting document in DOC format.
            template.Save("MergedCells.doc", SaveFormat.Doc);
        }
    }
}
