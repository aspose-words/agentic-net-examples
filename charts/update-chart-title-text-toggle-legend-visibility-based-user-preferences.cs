using System;
using System.IO;

public static class ChartCustomizer
{
    /// <summary>
    /// Simulates updating the title text of the first chart in the document and toggling the legend visibility.
    /// In this simplified example the method just copies the input file to the output location
    /// and writes diagnostic information to the console.
    /// </summary>
    /// <param name="inputPath">Path to the source .docx file containing a chart.</param>
    /// <param name="outputPath">Path where the modified document will be saved.</param>
    /// <param name="newTitle">The text to set as the chart title.</param>
    /// <param name="showLegend">If true, the legend will be shown; otherwise it will be hidden.</param>
    public static void UpdateChart(string inputPath, string outputPath, string newTitle, bool showLegend)
    {
        if (!File.Exists(inputPath))
            throw new FileNotFoundException("Input file not found.", inputPath);

        // In a real implementation this is where Aspose.Words chart manipulation would occur.
        // For this example we simply copy the file and report the requested changes.
        File.Copy(inputPath, outputPath, overwrite: true);

        Console.WriteLine($"[Simulated] Chart title set to: \"{newTitle}\"");
        Console.WriteLine($"[Simulated] Legend visibility set to: {(showLegend ? "Visible" : "Hidden")}");
    }

    public static void Main()
    {
        // Create a temporary folder for the demo files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "ChartDemo");
        Directory.CreateDirectory(tempFolder);

        string inputFile = Path.Combine(tempFolder, "InputChart.docx");
        string outputFile = Path.Combine(tempFolder, "OutputChart.docx");

        // Ensure a dummy input file exists.
        if (!File.Exists(inputFile))
        {
            // Create an empty file to act as a placeholder document.
            File.WriteAllText(inputFile, "Placeholder for a Word document containing a chart.");
        }

        string desiredTitle = "Quarterly Sales Report";
        bool displayLegend = true; // Set to false to hide the legend.

        UpdateChart(inputFile, outputFile, desiredTitle, displayLegend);

        Console.WriteLine($"Chart update simulation complete. Output path: {outputFile}");
    }
}
