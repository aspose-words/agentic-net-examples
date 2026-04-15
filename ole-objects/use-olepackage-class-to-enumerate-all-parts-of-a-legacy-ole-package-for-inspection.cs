using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Linq;

public class Program
{
    public static void Main()
    {
        // Create a temporary file to act as the source of the OLE package
        string tempFile = Path.Combine(Path.GetTempPath(), "SampleOlePackage.txt");
        File.WriteAllText(tempFile, "Hello, OLE Package!");

        // Locate and load the System.Windows.Forms assembly from the WindowsDesktop shared framework
        Assembly formsAssembly = LoadWindowsFormsAssembly();
        if (formsAssembly == null)
        {
            Console.WriteLine("Unable to locate System.Windows.Forms assembly.");
            Cleanup(tempFile);
            return;
        }

        // Get the OlePackage type
        Type olePackageType = formsAssembly.GetType("System.Windows.Forms.OlePackage");
        if (olePackageType == null)
        {
            Console.WriteLine("Unable to locate OlePackage type.");
            Cleanup(tempFile);
            return;
        }

        // Call OlePackage.FromFile(string) to create an instance
        MethodInfo fromFileMethod = olePackageType.GetMethod(
            "FromFile",
            BindingFlags.Public | BindingFlags.Static,
            null,
            new[] { typeof(string) },
            null);

        if (fromFileMethod == null)
        {
            Console.WriteLine("FromFile method not found.");
            Cleanup(tempFile);
            return;
        }

        object olePackage = fromFileMethod.Invoke(null, new object[] { tempFile });

        // Retrieve the GetDataFormats method (inherited from DataObject)
        MethodInfo getFormatsMethod = olePackageType.GetMethod(
            "GetDataFormats",
            BindingFlags.Public | BindingFlags.Instance,
            null,
            Type.EmptyTypes,
            null);

        if (getFormatsMethod == null)
        {
            Console.WriteLine("GetDataFormats method not found.");
            Cleanup(tempFile);
            return;
        }

        // Invoke GetDataFormats to obtain the list of formats (parts)
        string[] formats = (string[])getFormatsMethod.Invoke(olePackage, null);

        // Enumerate and display each format
        foreach (string format in formats)
        {
            Console.WriteLine($"Format: {format}");
        }

        // Clean up the temporary file
        Cleanup(tempFile);
    }

    private static Assembly LoadWindowsFormsAssembly()
    {
        // Get the directory of the core runtime (e.g., ...\shared\Microsoft.NETCore.App\8.0.0)
        string coreRuntimeDir = RuntimeEnvironment.GetRuntimeDirectory();

        // Move up to the shared folder and locate the WindowsDesktop shared framework
        string sharedDir = Path.GetFullPath(Path.Combine(coreRuntimeDir, "..", "Microsoft.WindowsDesktop.App"));
        if (!Directory.Exists(sharedDir))
            return null;

        // Choose the highest version folder (e.g., 8.0.0)
        string[] versionDirs = Directory.GetDirectories(sharedDir);
        if (versionDirs.Length == 0)
            return null;

        string latestVersionDir = versionDirs
            .OrderByDescending(v => v, StringComparer.OrdinalIgnoreCase)
            .First();

        // Build the full path to System.Windows.Forms.dll
        string assemblyPath = Path.Combine(latestVersionDir, "System.Windows.Forms.dll");
        if (!File.Exists(assemblyPath))
            return null;

        // Load the assembly from the discovered path
        return Assembly.LoadFrom(assemblyPath);
    }

    private static void Cleanup(string filePath)
    {
        try
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
        catch
        {
            // Ignored – cleanup failure shouldn't crash the program
        }
    }
}
