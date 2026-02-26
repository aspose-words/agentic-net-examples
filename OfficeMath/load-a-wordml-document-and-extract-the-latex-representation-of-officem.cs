using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Words.Math; // Added namespace for OfficeMath

public class OfficeMathLatexExtractor
{
    /// <summary>
    /// Loads a WORDML document and returns the LaTeX representation of each OfficeMath equation it contains.
    /// </summary>
    /// <param name="wordmlPath">Full path to the WORDML (.xml) document.</param>
    /// <returns>Array of LaTeX strings, one for each OfficeMath object found.</returns>
    public static string[] ExtractLatex(string wordmlPath)
    {
        // Ensure that shapes containing EquationXML are converted to OfficeMath objects during loading.
        LoadOptions loadOptions = new LoadOptions
        {
            ConvertShapeToOfficeMath = true
        };

        // Load the WORDML document.
        Document doc = new Document(wordmlPath, loadOptions);

        // Prepare save options that export OfficeMath as LaTeX.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            OfficeMathExportMode = TxtOfficeMathExportMode.Latex
        };

        // Collect LaTeX strings from all OfficeMath nodes.
        List<string> latexList = new List<string>();
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath math in mathNodes)
        {
            // Export the OfficeMath node to a LaTeX string using the configured options.
            string latex = math.ToString(txtOptions);
            latexList.Add(latex);
        }

        return latexList.ToArray();
    }

    // Example usage.
    public static void Main()
    {
        string wordmlFile = @"C:\Docs\SampleWordml.xml";

        string[] latexEquations = ExtractLatex(wordmlFile);

        Console.WriteLine("Extracted LaTeX equations:");
        foreach (string eq in latexEquations)
        {
            Console.WriteLine(eq);
        }
    }
}
