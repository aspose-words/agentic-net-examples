using System;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Saving;

class HyphenationPerformanceDemo
{
    static void Main()
    {
        // Create a new document and add a long paragraph to trigger hyphenation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 12;
        builder.Writeln(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut " +
            "labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco " +
            "laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in " +
            "voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat " +
            "non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.");

        // Prepare PDF save options (default options are sufficient for this demo).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // -----------------------------------------------------------------
        // Measure rendering time without hyphenation (default is false).
        // -----------------------------------------------------------------
        doc.HyphenationOptions.AutoHyphenation = false; // ensure hyphenation is disabled

        Stopwatch swNoHyphen = Stopwatch.StartNew();
        doc.Save("Hyphenation_Disabled.pdf", pdfOptions);
        swNoHyphen.Stop();

        Console.WriteLine($"PDF saved without hyphenation in {swNoHyphen.ElapsedMilliseconds} ms.");

        // -----------------------------------------------------------------
        // Measure rendering time with hyphenation enabled.
        // -----------------------------------------------------------------
        doc.HyphenationOptions.AutoHyphenation = true; // enable automatic hyphenation
        // Optional: fine‑tune hyphenation parameters.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

        Stopwatch swHyphen = Stopwatch.StartNew();
        doc.Save("Hyphenation_Enabled.pdf", pdfOptions);
        swHyphen.Stop();

        Console.WriteLine($"PDF saved with hyphenation in {swHyphen.ElapsedMilliseconds} ms.");

        // Output the difference.
        long diff = swHyphen.ElapsedMilliseconds - swNoHyphen.ElapsedMilliseconds;
        Console.WriteLine($"Hyphenation added {diff} ms to the rendering process.");
    }
}
