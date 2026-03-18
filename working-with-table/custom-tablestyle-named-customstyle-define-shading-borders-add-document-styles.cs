using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add a custom table style named "CustomStyle".
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "CustomStyle");

        // Define the shading (background color) for the style.
        customStyle.Shading.BackgroundPatternColor = Color.LightYellow;

        // Define the default borders for the style.
        customStyle.Borders.Color = Color.DarkBlue;
        customStyle.Borders.LineStyle = LineStyle.Single;
        customStyle.Borders.LineWidth = 1.0; // 1 point

        // Optionally, set additional formatting such as cell padding.
        customStyle.TopPadding = 5;
        customStyle.BottomPadding = 5;
        customStyle.LeftPadding = 5;
        customStyle.RightPadding = 5;

        // Save the document with the custom style.
        doc.Save("CustomTableStyle.docx");
    }
}
