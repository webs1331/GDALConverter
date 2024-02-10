using System.IO.Compression;
using OSGeo.OGR;

namespace KMZtoSHP;

internal class Program
{
    public static string PreviouslyRunFilesName { get; set; } = "PreviouslyConvertedFiles.txt";
    public static string WorkingDirectory { get; set; }
    public static List<string> PreviouslyConvertedFiles { get; set; } = new(); // DO NOT CHANGE THE NAMING CONVENTION - USED FOR SKIPPING PREVIOUSLY RUN FILE NAMES
    public static List<string> ConversionErrors { get; set; } = new();

    /// <summary>
    /// Convert a KMZ to SHP file using GDAL
    /// </summary>
    /// <param name="args"></param>
    public static void Main(string[] args)
    {
        var kmzFolderPath = PromptForKMZFolder();
        var shpFolderPath = PromptForSHPFolder();
        WorkingDirectory = $"{kmzFolderPath}\\temp\\";

        try
        {
            InitializeGDAL();
            ReadPreviouslyRunFilesMetadata(shpFolderPath);
            ConvertDirectory(kmzFolderPath, shpFolderPath);
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }
        finally
        {
            CleanUpWorkingDirectory();
            WritePreviouslyRunLog(shpFolderPath);
            WriteErrors(shpFolderPath);
        }
    }

    /// <summary>
    /// Converts all KMZ files in a directory to their respective (4) ESRI SHP file parts
    /// </summary>
    /// <param name="inputFolder">Folder location of the KMZ files to convert</param>
    /// <param name="outputFolder">Folder location of the SHP file outputs, and the log file outputs</param>
    private static void ConvertDirectory(string inputFolder, string outputFolder)
    {
        using Driver kmlDriver = Ogr.GetDriverByName("KML");
        using Driver esriShapeDriver = Ogr.GetDriverByName("ESRI Shapefile");

        var i = 1;
        var successfulConversions = 0;

        var KMZs = Directory.GetFiles(inputFolder, "*.kmz", SearchOption.AllDirectories);

        foreach (var inputKMZFile in KMZs)
        {
            try
            {
                if (PreviouslyConvertedFiles.Contains(inputKMZFile))
                {
                    Console.WriteLine($"Skipping {i} / {KMZs.Length} - Already converted");
                    continue;
                }

                //---------------Extract KMZ file to temp folder (should extract a single .KML file at this point)-------------
                ZipFile.ExtractToDirectory(inputKMZFile, WorkingDirectory, overwriteFiles: true);

                var getExtractedKMLFileName = Directory.GetFiles(WorkingDirectory, "*.kml").Single();
                var workingFileName = Path.Combine($"{inputFolder}\\temp\\", getExtractedKMLFileName);

                //----------------Read input KML file ----------------
                using DataSource kmlData = Ogr.Open(workingFileName, 0);
                using Layer kmlSelectedLayer = kmlData.GetLayerByIndex(0);

                //----------------Write output shp files ----------------
                var inputFileNameNoExt = Path.GetFileNameWithoutExtension(inputKMZFile);
                Directory.CreateDirectory($"{outputFolder}\\{inputFileNameNoExt}");
                var shapeData = esriShapeDriver.CreateDataSource($"{outputFolder}\\{inputFileNameNoExt}", options: new string[] { });
                using Layer newLayer = shapeData.CopyLayer(kmlSelectedLayer, inputFileNameNoExt.Replace("GIS", "").Replace("gis", "").TrimEnd(), new string[] { });

                Console.WriteLine($"Converted {successfulConversions++} / {KMZs.Length}");

                PreviouslyConvertedFiles.Add($"{inputKMZFile}");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Failed on {i} / {KMZs.Length}");
                Console.WriteLine($"Error converting {inputKMZFile}: {e.Message}");

                ConversionErrors.Add(inputKMZFile);
            }
            finally
            {
                i++;
                CleanUpWorkingDirectory();
            }
        }

        Console.WriteLine("Conversions complete");
        Console.WriteLine($"Successfully converted {successfulConversions} / {KMZs.Length}");
    }

    #region Setup and cleanup

    private static void InitializeGDAL()
    {
        GdalConfiguration.ConfigureGdal();
        GdalConfiguration.ConfigureOgr();
        Ogr.RegisterAll();
    }

    private static void CleanUpWorkingDirectory()
    {
        if (Directory.Exists(WorkingDirectory))
            Directory.Delete(WorkingDirectory, true);
    }

    //TODO - enter your own paths here
    private static string PromptForKMZFolder()
    {
        return "C:\\Users\\cqb13\\Desktop\\GIS KMZs";

        //Console.WriteLine("Enter the path to the folder containing the KMZ files: ");
        //return Console.ReadLine()!;
    }

    //TODO - enter your own paths here
    private static string PromptForSHPFolder()
    {
        return "C:\\Users\\cqb13\\Desktop\\GIS SHPs";

        //Console.WriteLine("Enter the path to the folder to save the SHP files: ");
        //return Console.ReadLine()!;
    }

    #endregion

    #region Log and config file interaction

    private static void ReadPreviouslyRunFilesMetadata(string outputPath)
    {
        var path = $"{outputPath}\\{PreviouslyRunFilesName}";

        if (File.Exists(path))
        {
            var lines = File.ReadAllLines(path);
            PreviouslyConvertedFiles.AddRange(lines);
        }
    }

    private static void WritePreviouslyRunLog(string outputPath) => File.WriteAllLines($"{outputPath}\\{PreviouslyRunFilesName}", PreviouslyConvertedFiles);

    private static void WriteErrors(string path)
    {
        if (ConversionErrors.Count > 0)
        {
            Console.WriteLine("---------------------------------------");
            Console.WriteLine("The following files failed to convert:");
            Console.WriteLine("---------------------------------------");

            foreach (var error in ConversionErrors)
                Console.WriteLine(error);

            File.WriteAllLines($"{path}//ConversionErrors-{DateTime.Now:yyyyMMdd-HHmmss}.txt", ConversionErrors);
        }
    }

    #endregion

    //hardcoded run single file test
    //public static void Convert()
    //{

    //    var fileName = "20-006 GIS";
    //    var kmlPath = @$"C:\Users\cqb13\Desktop\working folder\{fileName}";
    //    var shapeOutputPath = @"C:\Users\cqb13\Desktop\working folder\";

    //    ZipFile.ExtractToDirectory($"{kmlPath}.kmz", shapeOutputPath);

    //    using Driver kmlDriver = Ogr.GetDriverByName("KML");
    //    using DataSource kmlData = Ogr.Open($"{kmlPath}.kml", 0);
    //    using Layer kmlSelectedLayer = kmlData.GetLayerByIndex(0);
    //    using Driver esriShapeDriver = Ogr.GetDriverByName("ESRI Shapefile");

    //    var shapeData = esriShapeDriver.CreateDataSource(shapeOutputPath, new string[] { });

    //    using Layer newLayer = shapeData.CopyLayer(kmlSelectedLayer, fileName.Replace("GIS", "").Replace("gis", "").TrimEnd(), new string[] { });
    //}
}