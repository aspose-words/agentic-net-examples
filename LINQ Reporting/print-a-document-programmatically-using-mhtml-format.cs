using System;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Saving;

class PrintMhtmlExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Save the document in MHTML format.
        string mhtmlPath = @"C:\Temp\HelloWorld.mht";
        doc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Print the saved MHTML file using the default printer.
        // This uses the OS file association for the .mht extension.
        Process printProcess = new Process();
        printProcess.StartInfo.FileName = mhtmlPath;   // File to print.
        printProcess.StartInfo.Verb = "print";         // Use the print verb.
        printProcess.StartInfo.CreateNoWindow = true; // No console window.
        printProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
        printProcess.Start();

        // Optionally wait for the print job to be sent.
        printProcess.WaitForExit();
    }
}
