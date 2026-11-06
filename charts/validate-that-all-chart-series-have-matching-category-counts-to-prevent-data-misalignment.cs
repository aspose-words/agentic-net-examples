using System;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        if (!chartShape.HasChart)
            throw new InvalidOperationException("Inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Define categories for the chart using reflection (CategoryLabels may not exist in older versions).
        SetCategoryLabels(chart, new[] { "Q1", "Q2", "Q3" });

        // Remove any default series.
        chart.Series.Clear();

        // Add series with matching category counts.
        chart.Series.Add("Series 1", new double[] { 10, 20, 30 });
        chart.Series.Add("Series 2", new double[] { 15, 25, 35 });

        // Add a series with a mismatched number of values to demonstrate validation.
        chart.Series.Add("Series 3", new double[] { 5, 12 }); // Only 2 values instead of 3.

        // Validate that each series aligns with the category count.
        try
        {
            ValidateSeriesCategoryAlignment(chart);
            Console.WriteLine("All series have matching category counts.");
        }
        catch (InvalidOperationException ex)
        {
            Console.WriteLine($"Validation error: {ex.Message}");
        }

        // Save the document.
        doc.Save("validated-chart.docx");
    }

    // Sets the CategoryLabels property via reflection if it exists.
    private static void SetCategoryLabels(Chart chart, string[] labels)
    {
        PropertyInfo prop = chart.GetType().GetProperty("CategoryLabels", BindingFlags.Public | BindingFlags.Instance);
        if (prop != null && prop.CanWrite)
        {
            prop.SetValue(chart, labels);
        }
        else
        {
            // If the property does not exist, attempt to set XValues on the first series as a fallback.
            if (chart.Series.Count > 0)
            {
                ChartSeries firstSeries = chart.Series[0];
                PropertyInfo xValuesProp = firstSeries.GetType().GetProperty("XValues", BindingFlags.Public | BindingFlags.Instance);
                if (xValuesProp != null && xValuesProp.CanWrite)
                {
                    xValuesProp.SetValue(firstSeries, labels);
                }
            }
        }
    }

    private static void ValidateSeriesCategoryAlignment(Chart chart)
    {
        if (chart == null) throw new ArgumentNullException(nameof(chart));

        // Try to obtain category count from CategoryLabels property.
        int categoryCount = 0;
        PropertyInfo catProp = chart.GetType().GetProperty("CategoryLabels", BindingFlags.Public | BindingFlags.Instance);
        if (catProp != null)
        {
            var catArray = catProp.GetValue(chart) as string[];
            categoryCount = catArray?.Length ?? 0;
        }

        // Fallback: use the length of the first series' values if categories are not explicitly defined.
        if (categoryCount == 0 && chart.Series.Count > 0)
        {
            ChartSeries firstSeries = chart.Series[0];
            double[] values = GetSeriesValues(firstSeries);
            categoryCount = values?.Length ?? 0;
        }

        if (categoryCount == 0)
            throw new InvalidOperationException("Chart must contain at least one category label or series with values.");

        foreach (ChartSeries series in chart.Series)
        {
            double[] values = GetSeriesValues(series);
            if (values == null)
                throw new InvalidOperationException($"Series '{series.Name}' has no values.");

            if (values.Length != categoryCount)
                throw new InvalidOperationException(
                    $"Series '{series.Name}' has {values.Length} values, but there are {categoryCount} categories.");
        }
    }

    // Retrieves the Values array from a ChartSeries using reflection.
    private static double[] GetSeriesValues(ChartSeries series)
    {
        PropertyInfo valuesProp = series.GetType().GetProperty("Values", BindingFlags.Public | BindingFlags.Instance);
        if (valuesProp != null)
        {
            return valuesProp.GetValue(series) as double[];
        }

        // Some older versions expose the data via a method named GetValues().
        MethodInfo getValuesMethod = series.GetType().GetMethod("GetValues", BindingFlags.Public | BindingFlags.Instance);
        if (getValuesMethod != null)
        {
            return getValuesMethod.Invoke(series, null) as double[];
        }

        return null;
    }
}
