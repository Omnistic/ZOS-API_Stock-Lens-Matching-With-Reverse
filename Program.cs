using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using ZOSAPI;
using ZOSAPI.Editors.LDE;
using ZOSAPI.Tools.General;
using ZOSAPI.Editors;
using ZOSAPI.Tools.Optimization;
using ZOSAPI.Tools;

namespace CSharpUserExtensionApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            // Find the installed version of OpticStudio
            bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize();
            // Note -- uncomment the following line to use a custom initialization path
            //bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(@"C:\Program Files\OpticStudio\");
            if (isInitialized)
            {
                LogInfo("Found OpticStudio at: " + ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory());
            }
            else
            {
                HandleError("Failed to locate OpticStudio!");
                return;
            }
            
            BeginUserExtension();
        }

        static void BeginUserExtension()
        {
            // Create the initial connection class
            ZOSAPI_Connection TheConnection = new ZOSAPI_Connection();

            // Attempt to connect to the existing OpticStudio instance
            IZOSAPI_Application TheApplication = null;
            try
            {
                TheApplication = TheConnection.ConnectToApplication(); // this will throw an exception if not launched from OpticStudio
            }
            catch (Exception ex)
            {
                HandleError(ex.Message);
                return;
            }
            if (TheApplication == null)
            {
                HandleError("An unknown connection error occurred!");
                return;
            }
            if (TheApplication.Mode != ZOSAPI_Mode.Plugin)
            {
                HandleError("User plugin was started in the wrong mode: expected Plugin, found " + TheApplication.Mode.ToString());
                return;
            }
			
            // Chech the connection status
            if (!TheApplication.IsValidLicenseForAPI)
            {
                HandleError("Failed to connect to OpticStudio: " + TheApplication.LicenseStatus);
                return;
            }

            TheApplication.ProgressPercent = 0;
            TheApplication.ProgressMessage = "Reading settings ...";

            IOpticalSystem TheSystem = TheApplication.PrimarySystem;
			if (!TheApplication.TerminateRequested) // This will be 'true' if the user clicks on the Cancel button
            {
                string current_line, vendors_string, current_material, temporary_progress_message;
                double focal_length, epd, thickness_after, nominal_mf, current_mf, max_wavelength, min_wavelength, current_wavelength;
                int match_count, total_count, number_of_wavelengths;
                List<int> valid_optimization_cycles = new List<int> { 0, 1, 5, 10, 50 };
                ILensCatalogs TheLensCatalog;
                IMaterialsCatalog TheMaterialsCatalog;
                ILensCatalogLens MatchedLens;
                ISolveData thickness_solve;
                bool combinations = true;

                // Settings for catalog lens insertion
                bool ignore_object = true;
                bool reverse_geometry = true;

                // Default settings (in case the settings text file isn't found)
                bool reverse = true;
                bool variable = false;
                string[] vendors = new string[] { "EDMUND OPTICS", "THORLABS" };
                int matches = 5;
                double efl_tolerance = 0.25;
                double epd_tolerance = 0.25;
                bool air_compensation = true;
                int optimization_cycles = 0;
                bool save_best = true;

                // Data directory
                string data_directory = TheApplication.ZemaxDataDir;

                // Path to text-file settings
                string[] zosapi_folder = { data_directory, "ZOS-API", "Extensions", "Reverse_SLM_settings.txt" };
                string settings_path = Path.Combine(zosapi_folder);

                // Path to text-file log
                string[] log_folder = { data_directory, "ZOS-API", "Extensions", "Reverse_SLM_log.txt" };
                string log_path = Path.Combine(log_folder);

                // Output console log to text file (took from https://stackoverflow.com/questions/4470700/how-to-save-console-writeline-output-to-text-file/4470748)
                FileStream ostrm;
                StreamWriter writer;
                TextWriter oldOut = Console.Out;
                try
                {
                    ostrm = new FileStream(log_path, FileMode.Create, FileAccess.Write);
                    writer = new StreamWriter(ostrm);
                }
                catch (Exception e)
                {
                    Console.WriteLine("> ERROR: Cannot create a log-file, terminating user-extension ... ");
                    Console.WriteLine(e.Message);
                    return;
                }
                Console.SetOut(writer);

                // Try reading the settings from the text file
                try
                {
                    using (StreamReader sr = new StreamReader(settings_path))
                    {
                        reverse = sr.ReadLine().Contains("True") ? true : false;
                        variable = sr.ReadLine().Contains("True") ? true : false;
                        current_line = sr.ReadLine();
                        vendors_string = Regex.Replace(current_line.Substring(current_line.IndexOf("=") + 2), " *, *", ",");
                        vendors = vendors_string.Split(',');
                        current_line = sr.ReadLine();
                        matches = int.Parse(current_line.Substring(current_line.IndexOf("=") + 2));
                        current_line = sr.ReadLine();
                        efl_tolerance = double.Parse(current_line.Substring(current_line.IndexOf("=") + 2));
                        current_line = sr.ReadLine();
                        epd_tolerance = double.Parse(current_line.Substring(current_line.IndexOf("=") + 2));
                        air_compensation = sr.ReadLine().Contains("True") ? true : false;
                        current_line = sr.ReadLine();
                        optimization_cycles = int.Parse(current_line.Substring(current_line.IndexOf("=") + 2));
                        if (valid_optimization_cycles.IndexOf(optimization_cycles) == -1)
                        {
                            optimization_cycles = 0;
                            throw new Exception("Optimization cycles should be between 0 (automatic), 1, 5, 10, or 50");
                        }
                        save_best = sr.ReadLine().Contains("True") ? true : false;
                    }

                    LogInfo("> Settings read from text file ...");
                }
                // Otherwise, create a default settings text file
                catch
                {
                    // If the settings text file already exists, delete it
                    if (File.Exists(settings_path))
                    {
                        File.Delete(settings_path);
                    }

                    using (StreamWriter sw = File.CreateText(settings_path))
                    {
                        sw.WriteLine("Reverse = {0}", reverse);
                        sw.WriteLine("Variable = {0}", variable);
                        sw.WriteLine("Vendors = {0}", string.Join(",", vendors));
                        sw.WriteLine("Matches = {0}", matches);
                        sw.WriteLine("EFL tolerance = {0}", efl_tolerance);
                        sw.WriteLine("EPD tolerance = {0}", epd_tolerance);
                        sw.WriteLine("Air thickness compensation = {0}", air_compensation);
                        sw.WriteLine("Optimization cycles = {0}", optimization_cycles);
                        sw.WriteLine("Save best = {0}", save_best);
                    }

                    LogInfo("> WARNING: The settings text-file wasn't found or was incomplete, a default one has been created ...");
                }

                // Display current settings in console
                DisplaySettingsOnConsole(settings_path);

                // Update progress
                TheApplication.ProgressPercent = 5;
                TheApplication.ProgressMessage = "Finding lenses ...";
                Console.WriteLine("> Finding lenses ...");
                Console.WriteLine();

                // Retrieve maximum, and minimum wavelengths (used later to check if a matched lens material is compatible)
                max_wavelength = double.NegativeInfinity;
                min_wavelength = double.PositiveInfinity;
                number_of_wavelengths = TheSystem.SystemData.Wavelengths.NumberOfWavelengths;

                for (int ii = 1; ii <= number_of_wavelengths; ii++)
                {
                    current_wavelength = TheSystem.SystemData.Wavelengths.GetWavelength(ii).Wavelength;

                    if (current_wavelength > max_wavelength)
                    {
                        max_wavelength = current_wavelength;
                    }

                    if (current_wavelength < min_wavelength)
                    {
                        min_wavelength = TheSystem.SystemData.Wavelengths.GetWavelength(ii).Wavelength;
                    }
                }
                
                // Lens data editor
                ILensDataEditor TheLDE = TheSystem.LDE;

                // First surface to start the scan
                int first_surface = 1;

                // Last surface to scan
                int last_surface = TheLDE.NumberOfSurfaces;

                // Flags and counters to specify previous materials
                string previous_material = "";
                double semi_epd = -1;
                int element_count = 0;
                int lens_start = -1;
                int lens_count = 0;
                bool skipped = false;
                bool has_variable = false;

                // Maximum number of lenses to search a match for
                const int MAX_LENSES = 255;

                // Maximum number of elements for a single lens
                const int MAX_ELEMENTS = 3;

                // Array containing nominal lenses information
                object[][] lenses = new object[MAX_LENSES][];

                // Loop over the system surfaces
                for (int surface_number = first_surface; surface_number < last_surface; surface_number++)
                {
                    // Retrieve current surface
                    ILDERow current_surface = TheLDE.GetSurfaceAt(surface_number);

                    // We don't care about the material of the last surface
                    if (surface_number < last_surface)
                    {
                        // Retrieve current material
                        try
                        {
                            current_material = current_surface.Material;
                        }
                        catch
                        {
                            // Current material is air (returns a null object, which triggers an exception)
                            // Note: this also means the application doesn't treat Model Glass solves or
                            // any sort of solve for that matter
                            current_material = "";
                        }

                        // Were we in air?
                        if (previous_material == "")
                        {
                            // Are we still in air?
                            if (current_material == "")
                            {
                                // Then skip this surface
                                continue;
                            }
                            // Is it a mirror?
                            else if (current_material == "MIRROR")
                            {
                                // Then skip this surface as well
                                continue;
                            }
                            // Otherwise start a new lens
                            else
                            {
                                // Set new lens start surface
                                lens_start = surface_number;

                                // Semi EPD is the new surface clear semi-diameter
                                semi_epd = current_surface.SemiDiameter;

                                // Check if surface has variable radius
                                if (current_surface.RadiusCell.GetSolveData().Type == ZOSAPI.Editors.SolveType.Variable)
                                {
                                    has_variable = true;
                                }

                                // Update counter
                                element_count = 1;
                            }
                        }
                        else
                        {
                            // Are we completing a lens, i.e. returning in air?
                            if (current_material == "")
                            {
                                // If it has more than 3 elements, skip the lens (unsupported)
                                if (element_count > MAX_ELEMENTS)
                                {
                                    // Update counter
                                    element_count = 0;
                                    skipped = true;
                                    continue;
                                }

                                // Does the final element has a variable radius?
                                if (current_surface.RadiusCell.GetSolveData().Type == ZOSAPI.Editors.SolveType.Variable)
                                {
                                    has_variable = true;
                                }

                                // If it doesn't have at least one variable radius and variable is true, skip the lens
                                if (variable && !has_variable)
                                {
                                    // Update counter
                                    element_count = 0;
                                    continue;
                                }

                                // Retrieve focal length
                                focal_length = TheSystem.MFE.GetOperandValue(ZOSAPI.Editors.MFE.MeritOperandType.EFLX,
                                    lens_start, surface_number, 0, 0, 0, 0, 0, 0);

                                // Update list of lenses
                                lenses[lens_count] = new object[4] { element_count, lens_start, focal_length, 2 * semi_epd };

                                // Update counters
                                element_count = 0;
                                lens_count++;
                            }
                            // Is it a new element of the lens (I'm not exactly sure how to treat two consecutive
                            // surfaces with the same material yet), that is not a mirror?
                            else if ((current_material != previous_material) && (current_material != "MIRROR"))
                            {
                                // Has the EPD increased?
                                if (current_surface.SemiDiameter > semi_epd)
                                {
                                    // If so, update the EPD
                                    semi_epd = current_surface.SemiDiameter;
                                }

                                // Does the new element has a variable radius?
                                if (current_surface.RadiusCell.GetSolveData().Type == ZOSAPI.Editors.SolveType.Variable)
                                {
                                    has_variable = true;
                                }

                                // Update counters
                                element_count++;
                            }
                            else
                            {
                                // Then skip this surface (this includes the case where a lens is directly followed
                                // by a mirror, I don't know why anyone would do that but just in case)
                                continue;
                            }
                        }

                        // Update previous material
                        previous_material = current_material;
                    }
                }

                // Report if lenses with invalid number of elements
                if (skipped)
                {
                    Console.WriteLine("> WARNING: At least one lens with more than {0} elements has been ignored (unsupported) ...", MAX_ELEMENTS);
                    Console.WriteLine();
                }

                // Stop if no lenses are found
                if (lens_count == 0)
                {
                    // Update progress
                    Console.WriteLine("> ERROR: No lenses found to be matched, terminating user-extension ...");

                    // Clean up
                    FinishUserExtension(TheApplication);

                    // Restore console status
                    Console.SetOut(oldOut);
                    writer.Close();
                    ostrm.Close();

                    return;
                }

                // Update progress
                Console.WriteLine("\tFound {0} lens(es) in file.", lens_count);
                Console.WriteLine();

                // Check if the first vendor is "ALL"
                if (vendors[0] == "ALL" || vendors[0] == "all" || vendors[0] == "All")
                {
                    // Path to stock lens vendor catalogs
                    string[] vendors_folder = { data_directory, "Stockcat" };
                    string vendors_path = Path.Combine(vendors_folder);

                    // Get all catalog paths
                    string[] vendor_files = Directory.GetFiles(vendors_path, "*.ZMF");

                    // Extract the catalog filename only (this is the vendor)
                    vendors = new string[vendor_files.Length];

                    for (int ii = 0; ii < vendor_files.Length; ii++)
                    {
                        vendors[ii] = Path.GetFileNameWithoutExtension(vendor_files[ii]);
                    }
                }

                // Update progress
                Console.WriteLine("> Matching lenses ...");
                Console.WriteLine();

                // Path to temporary file
                string[] temporary_folder = { data_directory, "DeleteMe.ZMX" };
                string temporary_path = Path.Combine(temporary_folder);

                // Get nominal MF
                nominal_mf = TheSystem.MFE.CalculateMeritFunction();

                Console.WriteLine("\tNominal MF\t{0}", nominal_mf);
                Console.WriteLine();

                // Save variable air thicknesses
                List<int> air_thicknesses = VariableAirThicknesses(TheLDE, last_surface);

                // Save the system
                TheSystem.Save();

                // Open a copy of the system
                IOpticalSystem TheSystemCopy = TheSystem.CopySystem();

                // Copy of lens data editor
                ILensDataEditor TheLDECopy = TheSystemCopy.LDE;

                // Remove all variables
                TheSystemCopy.Tools.RemoveAllVariables();

                // Create a thickness variable solve
                thickness_solve = TheLDECopy.GetSurfaceAt(0).ThicknessCell.CreateSolveType(ZOSAPI.Editors.SolveType.Variable);

                // Have only previous variable air thicknesses as new variables
                foreach(int surface_id in air_thicknesses)
                {
                    TheLDECopy.GetSurfaceAt(surface_id).ThicknessCell.SetSolveData(thickness_solve);
                }

                // Save the copy of the system with correct variables
                TheSystemCopy.SaveAs(temporary_path);

                // Flag erroneous insertions
                bool insertion_error = false;
                bool material_error = false;

                // Array of best matches for each nominal lens
                object[,,] best_matches = new object[lens_count, matches, 6];

                // Loop over the lenses to be matched
                for (int lens_id = 0; lens_id < lens_count; lens_id++)
                {
                    // Initialization
                    total_count = 0;

                    // Nominal lens | Match ID | Reverse flag | MF value | Name | Vendor
                    for (int ii = 0; ii < matches; ii++)
                    {
                        best_matches[lens_id, ii, 0] = lens_id;
                        best_matches[lens_id, ii, 1] = -1;
                        best_matches[lens_id, ii, 2] = false;
                        best_matches[lens_id, ii, 3] = double.PositiveInfinity;
                        best_matches[lens_id, ii, 4] = "";
                        best_matches[lens_id, ii, 5] = "";
                    }

                    // Current lens properties
                    element_count = (int) lenses[lens_id][0];
                    lens_start = (int)lenses[lens_id][1];
                    focal_length = (double) lenses[lens_id][2];
                    epd = (double) lenses[lens_id][3];

                    // Iterate over the vendors
                    for (int vendor_id = 0; vendor_id < vendors.Length; vendor_id++)
                    {
                        // Run the lens catalog tool once to get the number of matches
                        TheLensCatalog = TheSystemCopy.Tools.OpenLensCatalogs();

                        // Apply settings to catalog tool
                        // Note: there are no error trapping for most of those settings!
                        ApplyCatalogSettings(TheLensCatalog, element_count, focal_length, epd, efl_tolerance, epd_tolerance, vendors[vendor_id]);

                        // Run the lens catalog tool for the given vendor
                        TheLensCatalog.RunAndWaitForCompletion();

                        // Save the number of matches
                        match_count = TheLensCatalog.MatchingLenses;
                        total_count += match_count;

                        // Close the lens catalog tool
                        TheLensCatalog.Close();

                        // Loop over the matches
                        for (int match_id = 0; match_id < match_count; match_id++)
                        {
                            // Check if terminate was pressed
                            if (TheApplication.TerminateRequested)
                            {
                                FinishUserExtension(TheApplication);

                                // Restore console status
                                Console.SetOut(oldOut);
                                writer.Close();
                                ostrm.Close();

                                return;
                            }

                            // Run the lens catalog once again and for every match (it needs to be closed for the optimization tool to be opened)
                            // Feature request: can we have a field to search for a specific lens by its name?
                            TheLensCatalog = TheSystemCopy.Tools.OpenLensCatalogs();
                            ApplyCatalogSettings(TheLensCatalog, element_count, focal_length, epd, efl_tolerance, epd_tolerance, vendors[vendor_id]);
                            TheLensCatalog.RunAndWaitForCompletion();

                            // Retrieve the corresponding matched lens
                            MatchedLens = TheLensCatalog.GetResult(match_id);

                            // Save the air thickness before next lens
                            thickness_after = TheSystemCopy.LDE.GetSurfaceAt(lens_start + element_count).Thickness;
                            thickness_solve = TheSystemCopy.LDE.GetSurfaceAt(lens_start + element_count).ThicknessCell.GetSolveData();

                            // Remove ideal lens
                            TheSystemCopy.LDE.RemoveSurfacesAt(lens_start, element_count + 1);

                            // Insert matching lens
                            if (!MatchedLens.InsertLensSeq(lens_start, ignore_object, reverse_geometry))
                            {
                                insertion_error = true;
                            }

                            // Restore thickness
                            TheSystemCopy.LDE.GetSurfaceAt(lens_start + element_count).Thickness = thickness_after;
                            TheSystemCopy.LDE.GetSurfaceAt(lens_start + element_count).ThicknessCell.SetSolveData(thickness_solve);

                            // Close the lens catalog tool
                            TheLensCatalog.Close();

                            // Check that material is within wavelength bounds (has to be done after closing the lens catalog)
                            for (int material_id = 0; material_id < element_count; material_id++)
                            {
                                // Material of every element of the lens
                                current_material = TheSystemCopy.LDE.GetSurfaceAt(lens_start + material_id).Material;

                                // Open the material catalog
                                TheMaterialsCatalog = TheSystemCopy.Tools.OpenMaterialsCatalog();

                                if (!MaterialIsCompatible(TheMaterialsCatalog, current_material, max_wavelength, min_wavelength))
                                {
                                    total_count--;
                                    material_error = true;
                                }

                                // Close the material catalog
                                TheMaterialsCatalog.Close();
                            }

                            // Update progress
                            TheApplication.ProgressPercent = 10 + 60 * (lens_id*vendors.Length*match_count + vendor_id*match_count + match_id) / (lens_count*vendors.Length + vendors.Length*match_count + match_count);
                            temporary_progress_message = "Lens " + (lens_id + 1).ToString() + "/" + lens_count.ToString();
                            temporary_progress_message += " | Vendor " + (vendor_id + 1).ToString() + "/" + (vendors.Length).ToString();
                            temporary_progress_message += " | Match " + (match_id + 1).ToString() + "/" + (match_count).ToString();
                            TheApplication.ProgressMessage = temporary_progress_message;

                            if (!material_error)
                            {
                                if (air_compensation)
                                {
                                    // Run the local optimizer
                                    ILocalOptimization TheOptimizer = TheSystemCopy.Tools.OpenLocalOptimization();
                                    TheOptimizer.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares;
                                    switch (optimization_cycles)
                                    {
                                        case 1:
                                            TheOptimizer.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Fixed_1_Cycle;
                                            break;
                                        case 5:
                                            TheOptimizer.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Fixed_5_Cycles;
                                            break;
                                        case 10:
                                            TheOptimizer.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Fixed_10_Cycles;
                                            break;
                                        case 50:
                                            TheOptimizer.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Fixed_50_Cycles;
                                            break;
                                        default:
                                            TheOptimizer.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Automatic;
                                            break;
                                    }
                                    TheOptimizer.RunAndWaitForCompletion();

                                    // Retrieve the MF value
                                    current_mf = TheOptimizer.CurrentMeritFunction;

                                    // Is it a best match?
                                    IsBestMatch(best_matches, lens_id, match_id, false, current_mf, MatchedLens.LensName, MatchedLens.Vendor);

                                    // Lens reversal enabled?
                                    if (reverse)
                                    {
                                        // Reverse the matched lens
                                        TheLDECopy.RunTool_ReverseElements(lens_start, lens_start + element_count);

                                        // Re-run the optimizer
                                        TheOptimizer.RunAndWaitForCompletion();

                                        // Retrieve the MF value
                                        current_mf = TheOptimizer.CurrentMeritFunction;

                                        // Is it a best match?
                                        IsBestMatch(best_matches, lens_id, match_id, true, current_mf, MatchedLens.LensName, MatchedLens.Vendor);
                                    }

                                    // Close the optimizer
                                    TheOptimizer.Close();
                                }
                                else
                                {
                                    // Retrieve the MF value
                                    current_mf = TheSystemCopy.MFE.CalculateMeritFunction();

                                    // Is it a best match?
                                    IsBestMatch(best_matches, lens_id, match_id, false, current_mf, MatchedLens.LensName, MatchedLens.Vendor);

                                    // Lens reversal enabled?
                                    if (reverse)
                                    {
                                        // Reverse the matched lens
                                        TheLDECopy.RunTool_ReverseElements(lens_start, lens_start + element_count);

                                        // Retrieve the MF value
                                        current_mf = TheSystemCopy.MFE.CalculateMeritFunction();

                                        // Is it a best match?
                                        IsBestMatch(best_matches, lens_id, match_id, true, current_mf, MatchedLens.LensName, MatchedLens.Vendor);
                                    }
                                }
                            }

                            // Load the copy of the original system
                            TheSystemCopy.LoadFile(temporary_path, false);
                        }
                    }

                    // Update progress
                    Console.WriteLine("\tMatched {0} lens(es) for lens {1} (Surface {2} to {3}).", total_count, lens_id+1, lens_start, lens_start+element_count);
                    if (total_count != 0)
                    {
                        for (int ii = 0; ii < matches; ii++)
                        {
                            if ((string) best_matches[lens_id, ii, 4] != "")
                            {
                                Console.WriteLine("\t  {0}. Merit function = {1}\t{2} ({3})\t\t[reversed = {4}]", ii + 1, best_matches[lens_id, ii, 3], best_matches[lens_id, ii, 4], best_matches[lens_id, ii, 5], best_matches[lens_id, ii, 2]);
                            }
                        }
                    }
                    else
                    {
                        combinations = false;
                    }
                }

                // Update progress
                Console.WriteLine();

                if (insertion_error)
                {
                    Console.WriteLine("> WARNING: One or more match failed to insert in the system ...");
                    Console.WriteLine();
                }

                if (material_error)
                {
                    Console.WriteLine("> WARNING: One or more match had a Material incompatible with the current wavelengths defined in the system ...");
                    Console.WriteLine();
                }

                // If combinations
                if (combinations)
                {
                    // Update progress
                    Console.WriteLine("> Combining matched lenses ...");
                    Console.WriteLine();


                }
                else
                {
                    // Clean up
                    FinishUserExtension(TheApplication);

                    Console.WriteLine("> WARNING: At least one lens had no match, combinations can't be investigated. Terminating the user-extension ...");

                    // Restore console status
                    Console.SetOut(oldOut);
                    writer.Close();
                    ostrm.Close();
                    return;
                }

                // Delete temporary file
                File.Delete(temporary_path);
                File.Delete(temporary_path.Replace(".ZMX", ".ZDA"));

                // Restore console status
                Console.SetOut(oldOut);
                writer.Close();
                ostrm.Close();
            }
			
			// Clean up
            FinishUserExtension(TheApplication);
        }

        static bool MaterialIsCompatible(IMaterialsCatalog TheMaterialsCatalog, string current_material, double max_wavelength, double min_wavelength)
        {
            bool compatible = false;
            
            string[] AllCatalogs = TheMaterialsCatalog.GetAllCatalogs();
            string[] AllMaterials;

            foreach(string Catalog in AllCatalogs)
            {
                TheMaterialsCatalog.SelectedCatalog = Catalog;

                // The materials are updated if the catalog is changed, the materials catalog does not need to be Run()
                AllMaterials = TheMaterialsCatalog.GetAllMaterials();

                if (Array.IndexOf(AllMaterials, current_material) != -1)
                {
                    TheMaterialsCatalog.SelectedMaterial = current_material;

                    // Once again, after selecting the material, one has directly access to the max/min wavelengths without running the tool
                    if (max_wavelength <= TheMaterialsCatalog.MaximumWavelength && min_wavelength >= TheMaterialsCatalog.MinimumWavelength)
                    {
                        compatible = true;
                    }

                    break;
                }
            }

            return compatible;
        }

        static void IsBestMatch(object[,,] best_matches, int lens_id, int match_id, bool reverse, double mf_value, string lens_name, string vendor)
        {
            for (int ii = 0; ii < best_matches.GetLength(1); ii++)
            {
                if (mf_value < (double) best_matches[lens_id, ii, 3])
                {
                    // Offset previous results
                    for (int jj = 4; jj-ii > 1; jj--)
                    {
                        best_matches[lens_id, jj, 1] = best_matches[lens_id, jj-1, 1];
                        best_matches[lens_id, jj, 2] = best_matches[lens_id, jj-1, 2];
                        best_matches[lens_id, jj, 3] = best_matches[lens_id, jj-1, 3];
                        best_matches[lens_id, jj, 4] = best_matches[lens_id, jj-1, 4];
                        best_matches[lens_id, jj, 5] = best_matches[lens_id, jj-1, 5];
                    }

                    // Save new match as best
                    best_matches[lens_id, ii, 1] = match_id;
                    best_matches[lens_id, ii, 2] = reverse;
                    best_matches[lens_id, ii, 3] = mf_value;
                    best_matches[lens_id, ii, 4] = lens_name;
                    best_matches[lens_id, ii, 5] = vendor;

                    break;
                }
            }
        }

        static List<int> VariableAirThicknesses(ILensDataEditor TheLDE, int last_surface)
        {
            List<int> SurfaceIDs = new List<int>();

            for (int ii = 1; ii < last_surface; ii++)
            {
                if (TheLDE.GetSurfaceAt(ii).ThicknessCell.GetSolveData().Type == ZOSAPI.Editors.SolveType.Variable)
                {
                    SurfaceIDs.Add(ii);
                }
            }

            return SurfaceIDs;
        }

        static void ApplyCatalogSettings(ILensCatalogs TheLensCatalog, int element_count, double focal_length, double epd, double efl_tolerance, double epd_tolerance, string vendor)
        {
            // Vendor
            TheLensCatalog.SelectedVendor = vendor;

            // EFL
            TheLensCatalog.UseEFL = true;
            TheLensCatalog.MinEFL = focal_length - focal_length * efl_tolerance;
            TheLensCatalog.MaxEFL = focal_length + focal_length * efl_tolerance;

            // EPD
            TheLensCatalog.UseEPD = true;
            TheLensCatalog.MinEPD = epd - epd * epd_tolerance;
            TheLensCatalog.MaxEPD = epd + epd * epd_tolerance;

            // Supported lens type
            TheLensCatalog.IncShapeEqui = true;
            TheLensCatalog.IncShapeBi = true;
            TheLensCatalog.IncShapePlano = true;
            TheLensCatalog.IncShapeMeniscus = true;
            TheLensCatalog.IncTypeSpherical = true;

            // Unsupported lens type
            TheLensCatalog.IncTypeGRIN = false;
            TheLensCatalog.IncTypeAspheric = false;
            TheLensCatalog.IncTypeToroidal = false;

            // Number of elements in the lens
            TheLensCatalog.NumberOfElements = element_count;
        }

        static void DisplaySettingsOnConsole(string path)
        {
            Console.WriteLine();
            Console.WriteLine("=================== Stock Lens Matching Settings ===================");

            using (StreamReader sr = new StreamReader(path))
            {
                while (sr.Peek() >= 0)
                {
                    Console.Write((char)sr.Read());
                }
            }

            Console.WriteLine("=====================================================================");
            Console.WriteLine();
        }

		static void FinishUserExtension(IZOSAPI_Application TheApplication)
		{
            // Note - OpticStudio will stay in User Extension mode until this application exits
			if (TheApplication != null)
			{
                TheApplication.ProgressMessage = "Complete";
                TheApplication.ProgressPercent = 100;
			}
		}

        static void LogInfo(string message)
        {
            // TODO - add custom logging
            Console.WriteLine(message);
        }

        static void HandleError(string errorMessage)
        {
            // TODO - add custom error handling
            throw new Exception(errorMessage);
        }

    }
}
