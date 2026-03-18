using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

// Create a new document and a DocumentBuilder.
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

// Insert a column chart into the document.
Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
Chart chart = chartShape.Chart;

// Configure the plot area border: dashed line, red color, 2 pt width.
chart.Format.Stroke.Color = Color.Red;          // Border color.
chart.Format.Stroke.Weight = 2.0;               // Border width in points.
chart.Format.Stroke.DashStyle = DashStyle.Dash; // Dashed line style.

// Save the document.
doc.Save("PlotAreaBorder.docx");
