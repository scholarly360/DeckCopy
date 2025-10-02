# DeckCopy :: Utility to Copy Slides from Source to Target (Using C#)

This C# program is a PowerPoint Presentation Merger CLI that uses the OpenXML SDK to merge slides from one PowerPoint presentation into another. Here's how it works:

## Main Functionality
The program allows you to:

Copy specific slides or all slides from a source PowerPoint file
Insert them into a target PowerPoint file
Preserve the master slides and formatting from the target presentation
Output the result as a new merged presentation file

## Command Line Interface
The program accepts these arguments:

-s/--source: Source PowerPoint file (.pptx)
-t/--target: Target PowerPoint file (.pptx)
-o/--output: Output file (optional, defaults to target_merged.pptx)
--slides: Specific slides to copy (optional, copies all if not specified)
-h/--help: Show help information

##  Key Features
Flexible Slide Selection
Single slides: 1,3,5
Ranges: 2-5,8-10
Mixed: 1,3-5,7,9-12


## DEVELOPMENT :::

dotnet restore DeckCopy.csproj

dotnet build DeckCopy.csproj

dotnet run --project DeckCopy.csproj

---------------------------------------------------

## RUNNING AS CLI ::: 

Copy all slides:

dotnet run -- -s 002.pptx -t 001.pptx

Copy specific slides:

dotnet run -- -s 002.pptx -t 001.pptx --slides 2,3

Specify custom output file:

dotnet run -- -s 002.pptx -t 001.pptx -o merged.pptx --slides 1,2,3
---------------------------------------------------

## Publish as Self-Contained Executable :::

### Compile the project:
dotnet build DeckCopy.csproj --configuration Release

### For Windows (x64) Preferred
dotnet publish DeckCopy.csproj -c Release -r win-x64 --self-contained true -o ./publish/win-x64

### For Linux (x64) Optional for Future
dotnet publish DeckCopy.csproj -c Release -r linux-x64 --self-contained true -o ./publish/linux-x64

### For macOS (x64) Optional for Future
dotnet publish DeckCopy.csproj -c Release -r osx-x64 --self-contained true -o ./publish/osx-x64

## Single File Executable (Preferred)
dotnet publish DeckCopy.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o ./publish/PPTJoiner

## RUN as CLI
DeckCopy -s 002.pptx -t 001.pptx -o merged.pptx --slides 1,2,3 

## Run from Python (IMP)
import subprocess

cmd = [
    "DeckCopy",   # the exe (must be in PATH, else give full path)
    "-s", "002.pptx",
    "-t", "001.pptx",
    "-o", "merged.pptx",
    "--slides", "1,2,3"
]

result = subprocess.run(cmd, capture_output=True, text=True)

print("Exit Code:", result.returncode)
print("STDOUT:", result.stdout)
print("STDERR:", result.stderr)
