using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Math; // <-- required for OfficeMath class
using Aspose.Words.Saving;

class ExtractOfficeMathLatex
{
    static void Main()
    {
        // Path to the WORDML (WordprocessingML) document.
        string wordmlPath = @"C:\Docs\SampleWordML.xml";

        // Load the WORDML document.
        Document doc = new Document(wordmlPath);

        // Retrieve all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Prepare save options that export OfficeMath as LaTeX.
        TxtSaveOptions latexOptions = new TxtSaveOptions
        {
            OfficeMathExportMode = TxtOfficeMathExportMode.Latex
        };

        // Iterate through each OfficeMath node and extract its LaTeX representation.
        List<string> latexEquations = new List<string>();
        foreach (Node node in mathNodes)
        {
            // Cast the node to OfficeMath.
            OfficeMath math = (OfficeMath)node;

            // The ToString method with TxtSaveOptions returns the LaTeX string for this node.
            string latex = math.ToString(latexOptions);
            latexEquations.Add(latex);
        }

        // Output the extracted LaTeX equations.
        Console.WriteLine("Extracted LaTeX equations:");
        foreach (string eq in latexEquations)
        {
            Console.WriteLine(eq);
        }
    }
}
