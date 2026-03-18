using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Reporting;

class ReportWithChart
{
    static void Main()
    {
        // 1. Prepare the data source – a DataSet containing one DataTable.
        DataSet dataSet = new DataSet();

        DataTable salesTable = new DataTable("Sales");
        salesTable.Columns.Add("Year", typeof(int));
        salesTable.Columns.Add("Amount", typeof(double));

        // Sample numeric data.
        salesTable.Rows.Add(2018, 12500.75);
        salesTable.Rows.Add(2019, 15830.20);
        salesTable.Rows.Add(2020, 14200.00);
        salesTable.Rows.Add(2021, 17650.55);
        salesTable.Rows.Add(2022, 19000.10);

        dataSet.Tables.Add(salesTable);

        // 2. Create a blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 3. Insert a column chart that will later be filled with the data.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // 4. Clear the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // 5. Extract the numeric values from the DataTable.
        //    X‑axis – years, Y‑axis – amounts.
        int rowCount = salesTable.Rows.Count;
        double[] years = new double[rowCount];
        double[] amounts = new double[rowCount];

        for (int i = 0; i < rowCount; i++)
        {
            years[i] = Convert.ToDouble(salesTable.Rows[i]["Year"]);
            amounts[i] = Convert.ToDouble(salesTable.Rows[i]["Amount"]);
        }

        // 6. Add a series to the chart using the extracted arrays.
        chart.Series.Add("Annual Sales", years, amounts);

        // 7. (Optional) Show the data table beneath the chart.
        ChartDataTable dataTable = chart.DataTable;
        dataTable.Show = true;
        dataTable.HasLegendKeys = false;
        dataTable.HasHorizontalBorder = false;
        dataTable.HasVerticalBorder = false;
        dataTable.HasOutlineBorder = false;
        dataTable.Font.Italic = true;

        // 8. Use ReportingEngine to merge the DataSet into the document.
        //    In this simple scenario the chart is already populated, but the engine
        //    can also be used to replace other template fields if needed.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSet, "ds"); // "ds" is the name used in the template (if any).

        // 9. Save the final report.
        doc.Save("ReportWithChart.docx");
    }
}
