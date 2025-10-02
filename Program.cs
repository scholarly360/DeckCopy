/* This is CLI : Vibe Coded for Python Utiization */ 
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Collections.Generic;
using System.IO;
using System;
using System.Linq;

internal class Program
{
    // Ensure only one Main method exists in the project.
    private static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            ShowHelp();
            return;
        }

        try
        {
            ParseArguments(args);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine();
            ShowHelp();
            Environment.Exit(1);
        }
    }

    private static void ParseArguments(string[] args)
    {
        string sourceFile = null;
        string targetFile = null;
        string outputFile = null;
        List<int> copySlideNo = new List<int>();
        bool showHelp = false;

        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i].ToLower())
            {
                case "-s":
                case "--source":
                    if (i + 1 < args.Length)
                        sourceFile = args[++i];
                    else
                        throw new ArgumentException("Source file path is required after -s/--source");
                    break;

                case "-t":
                case "--target":
                    if (i + 1 < args.Length)
                        targetFile = args[++i];
                    else
                        throw new ArgumentException("Target file path is required after -t/--target");
                    break;

                case "-o":
                case "--output":
                    if (i + 1 < args.Length)
                        outputFile = args[++i];
                    else
                        throw new ArgumentException("Output file path is required after -o/--output");
                    break;

                case "--slides":
                    if (i + 1 < args.Length)
                    {
                        string slidesArg = args[++i];
                        copySlideNo = ParseSlideNumbers(slidesArg);
                    }
                    else
                        throw new ArgumentException("Slide numbers are required after --slides");
                    break;

                case "-h":
                case "--help":
                    showHelp = true;
                    break;

                default:
                    throw new ArgumentException($"Unknown argument: {args[i]}");
            }
        }

        if (showHelp)
        {
            ShowHelp();
            return;
        }

        // Validate required arguments
        if (string.IsNullOrEmpty(sourceFile))
            throw new ArgumentException("Source file is required. Use -s or --source");
        
        if (string.IsNullOrEmpty(targetFile))
            throw new ArgumentException("Target file is required. Use -t or --target");

        // Set default output file if not provided
        if (string.IsNullOrEmpty(outputFile))
        {
            string targetFileNameWithoutExt = Path.GetFileNameWithoutExtension(targetFile);
            string targetDir = Path.GetDirectoryName(targetFile);
            outputFile = Path.Combine(targetDir ?? "", $"{targetFileNameWithoutExt}_merged.pptx");
        }

        // Validate files exist
        if (!File.Exists(sourceFile))
            throw new FileNotFoundException($"Source file not found: {sourceFile}");
        
        if (!File.Exists(targetFile))
            throw new FileNotFoundException($"Target file not found: {targetFile}");

        // If no specific slides specified, copy all slides
        if (copySlideNo.Count == 0)
        {
            Console.WriteLine("No specific slides specified. Will copy all slides from source.");
            copySlideNo = GetAllSlideNumbers(sourceFile);
        }

        // Execute the merge
        ExecuteMerge(sourceFile, targetFile, outputFile, copySlideNo);
    }

    private static void ExecuteMerge(string sourceFile, string targetFile, string outputFile, List<int> copySlideNo)
    {
        try
        {
            Console.WriteLine("Starting selective presentation merge...");
            Console.WriteLine($"Source: {sourceFile}");
            Console.WriteLine($"Target: {targetFile}");
            Console.WriteLine($"Output: {outputFile}");
            Console.WriteLine($"Slides to copy: {string.Join(", ", copySlideNo)}");
            Console.WriteLine();
            
            // Create a copy of the target file to work with
            File.Copy(targetFile, outputFile, true);
            
            MergePresentations(sourceFile, outputFile, copySlideNo);
            
            Console.WriteLine($"Successfully merged presentations. Output saved as: {outputFile}");
            Console.WriteLine("The master slide from the target presentation has been preserved.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during merge: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"Inner exception: {ex.InnerException.Message}");
            throw;
        }
    }

    private static List<int> ParseSlideNumbers(string slidesArg)
    {
        var slideNumbers = new List<int>();
        var parts = slidesArg.Split(',');

        foreach (var part in parts)
        {
            var trimmedPart = part.Trim();
            
            if (trimmedPart.Contains('-'))
            {
                // Handle range like "3-7"
                var rangeParts = trimmedPart.Split('-');
                if (rangeParts.Length == 2 && 
                    int.TryParse(rangeParts[0].Trim(), out int start) && 
                    int.TryParse(rangeParts[1].Trim(), out int end))
                {
                    slideNumbers.AddRange(CreateSlideRange(start, end));
                }
                else
                {
                    throw new ArgumentException($"Invalid slide range format: {trimmedPart}. Use format like '3-7'");
                }
            }
            else
            {
                // Handle single slide number
                if (int.TryParse(trimmedPart, out int slideNum))
                {
                    slideNumbers.Add(slideNum);
                }
                else
                {
                    throw new ArgumentException($"Invalid slide number: {trimmedPart}");
                }
            }
        }

        return slideNumbers.Distinct().OrderBy(x => x).ToList();
    }

    private static List<int> GetAllSlideNumbers(string sourceFile)
    {
        using (PresentationDocument sourceDoc = PresentationDocument.Open(sourceFile, false))
        {
            var sourceSlideIds = sourceDoc.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();
            return Enumerable.Range(1, sourceSlideIds.Count).ToList();
        }
    }

    private static void ShowHelp()
    {
        Console.WriteLine("PowerPoint Presentation Merger CLI");
        Console.WriteLine("==================================");
        Console.WriteLine();
        Console.WriteLine("Usage: vsc_csharp_openxmlsdk [options]");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  -s, --source <file>     Source PowerPoint file (.pptx)");
        Console.WriteLine("  -t, --target <file>     Target PowerPoint file (.pptx) - master slides preserved");
        Console.WriteLine("  -o, --output <file>     Output file (optional, defaults to target_merged.pptx)");
        Console.WriteLine("  --slides <numbers>      Specific slides to copy (optional, defaults to all)");
        Console.WriteLine("  -h, --help              Show this help message");
        Console.WriteLine();
        Console.WriteLine("Slide number formats:");
        Console.WriteLine("  Single slides: 1,3,5");
        Console.WriteLine("  Ranges: 2-5,8-10");
        Console.WriteLine("  Mixed: 1,3-5,7,9-12");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  # Merge all slides from source to target");
        Console.WriteLine("  vsc_csharp_openxmlsdk -s source.pptx -t target.pptx");
        Console.WriteLine();
        Console.WriteLine("  # Merge specific slides");
        Console.WriteLine("  vsc_csharp_openxmlsdk -s source.pptx -t target.pptx --slides 1,3,5-7");
        Console.WriteLine();
        Console.WriteLine("  # Specify custom output file");
        Console.WriteLine("  vsc_csharp_openxmlsdk -s source.pptx -t target.pptx -o merged.pptx --slides 2-4");
    }

    private static void MergePresentations(string sourceFile, string targetFile, List<int> copySlideNo)
    {
        using (PresentationDocument sourceDoc = PresentationDocument.Open(sourceFile, false))
        using (PresentationDocument targetDoc = PresentationDocument.Open(targetFile, true))
        {
            // Get the presentation parts
            PresentationPart sourcePresentationPart = sourceDoc.PresentationPart;
            PresentationPart targetPresentationPart = targetDoc.PresentationPart;

            // Get the slide ID list from both presentations
            var sourceSlideIds = sourcePresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();
            var targetSlideIdList = targetPresentationPart.Presentation.SlideIdList;

            Console.WriteLine($"Source presentation has {sourceSlideIds.Count} slides");
            Console.WriteLine($"Target presentation has {targetSlideIdList.Elements<SlideId>().Count()} slides");

            // Validate slide numbers (convert from 1-based to 0-based indexing)
            var validSlideNumbers = copySlideNo.Where(n => n >= 1 && n <= sourceSlideIds.Count).ToList();
            var invalidSlideNumbers = copySlideNo.Where(n => n < 1 || n > sourceSlideIds.Count).ToList();

            if (invalidSlideNumbers.Any())
            {
                Console.WriteLine($"Warning: Invalid slide numbers will be skipped: {string.Join(", ", invalidSlideNumbers)}");
            }

            Console.WriteLine($"Will copy {validSlideNumbers.Count} slides: {string.Join(", ", validSlideNumbers)}");

            // Keep track of the maximum slide ID to avoid conflicts
            uint maxSlideId = 0;
            if (targetSlideIdList.Elements<SlideId>().Any())
            {
                maxSlideId = targetSlideIdList.Elements<SlideId>().Max(s => s.Id.Value);
            }

            // Get the first available slide layout from target to ensure consistency
            SlideLayoutPart targetLayout = null;
            if (targetPresentationPart.SlideMasterParts.Any())
            {
                targetLayout = targetPresentationPart.SlideMasterParts.First().SlideLayoutParts.FirstOrDefault();
            }

            // Copy only selected slides from source to target
            foreach (int slideNumber in validSlideNumbers)
            {
                // Convert from 1-based to 0-based indexing
                int slideIndex = slideNumber - 1;
                var sourceSlideId = sourceSlideIds[slideIndex];
                
                Console.WriteLine($"Copying slide {slideNumber} (index {slideIndex})...");
                
                // Get the source slide part
                SlidePart sourceSlide = (SlidePart)sourcePresentationPart.GetPartById(sourceSlideId.RelationshipId);
                
                // Create a new slide part in the target presentation
                SlidePart newSlidePart = targetPresentationPart.AddNewPart<SlidePart>();
                
                // Clone the slide content
                var clonedSlide = (Slide)sourceSlide.Slide.CloneNode(true);
                
                // Remove any layout references from the cloned slide to avoid conflicts
                var commonSlideData = clonedSlide.GetFirstChild<CommonSlideData>();
                if (commonSlideData != null)
                {
                    // Keep the slide content but ensure it works with target's master
                    newSlidePart.Slide = clonedSlide;
                    
                    // Associate with target's layout if available
                    if (targetLayout != null)
                    {
                        newSlidePart.AddPart(targetLayout);
                    }
                }
                else
                {
                    // Create a minimal slide if the source slide is malformed
                    newSlidePart.Slide = new Slide(
                        new CommonSlideData(
                            new ShapeTree(
                                new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                                    new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() { Id = 1, Name = "" },
                                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                                    new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()),
                                new DocumentFormat.OpenXml.Presentation.GroupShapeProperties(
                                    new DocumentFormat.OpenXml.Drawing.TransformGroup()))));
                    
                    if (targetLayout != null)
                    {
                        newSlidePart.AddPart(targetLayout);
                    }
                }
                
                // Create a new SlideId for the target presentation
                SlideId newSlideId = new SlideId()
                {
                    Id = ++maxSlideId,
                    RelationshipId = targetPresentationPart.GetIdOfPart(newSlidePart)
                };

                // Add the new slide ID to the target presentation
                targetSlideIdList.Append(newSlideId);
            }

            // Ensure the presentation structure is valid
            if (targetPresentationPart.Presentation.SlideSize == null)
            {
                targetPresentationPart.Presentation.SlideSize = new SlideSize() { Cx = 9144000, Cy = 6858000 };
            }

            // Save the changes
            targetPresentationPart.Presentation.Save();
            
            Console.WriteLine($"Successfully copied {validSlideNumbers.Count} selected slides from '{sourceFile}' to '{targetFile}'");
            Console.WriteLine($"Copied slides: {string.Join(", ", validSlideNumbers)}");
            Console.WriteLine("Master slides from the target presentation have been preserved.");
        }
    }

    // Overloaded method for copying all slides (maintains backward compatibility)
    private static void MergePresentations(string sourceFile, string targetFile)
    {
        using (PresentationDocument sourceDoc = PresentationDocument.Open(sourceFile, false))
        {
            var sourceSlideIds = sourceDoc.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();
            var allSlideNumbers = Enumerable.Range(1, sourceSlideIds.Count).ToList();
            MergePresentations(sourceFile, targetFile, allSlideNumbers);
        }
    }

    // Helper method to create slide range (e.g., slides 3-7)
    public static List<int> CreateSlideRange(int start, int end)
    {
        return Enumerable.Range(start, end - start + 1).ToList();
    }

    // Helper method to combine multiple slide selections
    public static List<int> CombineSlideSelections(params List<int>[] selections)
    {
        return selections.SelectMany(x => x).Distinct().OrderBy(x => x).ToList();
    }
}
