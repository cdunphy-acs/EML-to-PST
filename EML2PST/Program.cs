/* Taken from https://github.com/THJLI/eml.to.pst.outlook
 * Enahnced by Chris Dunphy
*/


using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using System.Net.Mail;
using ShellProgressBar;

var dirPath = "ENTER PATH HERE";
var outFileName = $"{dirPath}\\outputFile.pst";

if (File.Exists(outFileName))
    File.Delete(outFileName);

using (var personalStorage = PersonalStorage.Create(outFileName, FileFormatVersion.Unicode))
{
    var directories = Directory.GetDirectories(dirPath);
    if (directories.Any())
    {
        using (var pbar = new ProgressBar(directories.Length, "Processing directories...", new ProgressBarOptions { ProgressCharacter = '─', ProgressBarOnBottom = true }))
        {
            foreach (var d in directories)
            {
                SetFilesIntoBox(personalStorage, d);
                pbar.Tick($"Finished processing {d}");
            }
        }
    }
    else
        SetFilesIntoBox(personalStorage, dirPath);
}

void SetFilesIntoBox(PersonalStorage personalStorage, string directoryPath)
{
    var pathBox = personalStorage.RootFolder.AddSubFolder(Path.GetFileName(directoryPath));
    Console.WriteLine($"Create box: {pathBox.DisplayName}");
    var files = Directory.GetFiles(directoryPath, "*.eml");
    using (var pbar = new ProgressBar(files.Length, "Processing files...", new ProgressBarOptions { ProgressCharacter = '─', ProgressBarOnBottom = true }))
    {
        for (int i = 0; i < files.Length; i++)
        {
            var f = files[i];
            using (var message = Aspose.Email.MailMessage.Load(f))
            {
                pathBox.AddMessage(MapiMessage.FromMailMessage(message, MapiConversionOptions.UnicodeFormat));
                pbar.Tick($"Finished processing {f}. Processed: {i + 1}/{files.Length}");
            }
        }
    }
}
