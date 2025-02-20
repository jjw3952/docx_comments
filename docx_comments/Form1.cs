//using System;
//using System.Diagnostics;
//using System.IO;
//using System.Windows.Forms;

//public class MainForm : Form
//{
//    private Button selectDirectoryButton;
//    private TextBox outputBox;

//    public MainForm()
//    {
//        // Initialize the form
//        this.Text = "Select Directory to Extract Comments";
//        this.Size = new System.Drawing.Size(600, 400);

//        selectDirectoryButton = new Button { 
//            Text = "Select Files", //"Select Directory",
//            Dock = DockStyle.Top,
//            BackColor = System.Drawing.Color.Yellow
//        };
//        selectDirectoryButton.Click += SelectFiles; //SelectDirectory;

//        outputBox = new TextBox { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical };

//        Controls.Add(outputBox);
//        Controls.Add(selectDirectoryButton);
//    }

//    // These are commented out as they processed all DOCX files in a directory.
//    // Now I have it setup to process individually selected files.
//    //private void SelectDirectory(object sender, EventArgs e)
//    //{
//    //    using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
//    //    {
//    //        if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
//    //        {
//    //            string selectedPath = folderBrowserDialog.SelectedPath;
//    //            outputBox.Text = $"Selected Directory: {selectedPath}" + Environment.NewLine;
//    //            ProcessDocxFiles(selectedPath);
//    //        }
//    //    }
//    //}

//    //private void ProcessDocxFiles(string directoryPath)
//    //{
//    //    string[] docxFiles = Directory.GetFiles(directoryPath, "*.docx");
//    //    foreach (string docxFile in docxFiles)
//    //    {
//    //        RunRScript(docxFile, directoryPath);
//    //    }
//    //}

//    private void SelectFiles(object sender, EventArgs e)
//    {
//        using (OpenFileDialog openFileDialog = new OpenFileDialog())
//        {
//            openFileDialog.Multiselect = true;
//            openFileDialog.Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*";
//            if (openFileDialog.ShowDialog() == DialogResult.OK)
//            {
//                string[] selectedFiles = openFileDialog.FileNames;
//                outputBox.Text = "Selected Files:" + Environment.NewLine;
//                foreach (string file in selectedFiles)
//                {
//                    outputBox.Text += file + Environment.NewLine;
//                }
//                ProcessDocxFiles(selectedFiles);
//            }
//        }
//    }

//    private void ProcessDocxFiles(string[] filePaths)
//    {
//        foreach (string filePath in filePaths)
//        {
//            string directoryPath = Path.GetDirectoryName(filePath);
//            RunRScript(filePath, directoryPath);
//        }
//    }

//    private void RunRScript(string filePath, string outputDirectory)
//    {
//        ProcessStartInfo psi = new ProcessStartInfo
//        {
//            FileName = "C:\\Program Files\\R\\R-4.4.0\\bin\\Rscript.exe",
//            Arguments = $"\"C:\\Users\\AG_User_10\\source\\repos\\docx_comments\\extract_docx_comments.R\" \"{filePath}\" \"{outputDirectory}\"",
//            RedirectStandardOutput = true,
//            UseShellExecute = false,
//            CreateNoWindow = true
//        };

//        using (Process process = new Process { StartInfo = psi })
//        {
//            process.Start();
//            string output = process.StandardOutput.ReadToEnd();
//            string[] outputLines = output.Split(new[] { "!" }, StringSplitOptions.None);
//            foreach (string line in outputLines)
//            {
//                outputBox.Text += line + Environment.NewLine;
//            }
//        }
//    }
//    catch (Exception ex)
//    {
//        outputBox.Text += $"Error: {ex.Message}" + Environment.NewLine;
//    }
//}

//public static class Program
//{
//    [STAThread]
//    public static void Main()
//    {
//        Application.EnableVisualStyles();
//        Application.SetCompatibleTextRenderingDefault(false);
//        Application.Run(new MainForm());
//    }
//}

using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Linq; // Add this using directive

public class Config
{
    public string RscriptPath { get; set; }
    public string RScriptFilePath { get; set; }
}

public class MainForm : Form
{
    private Button selectDirectoryButton;
    private TextBox outputBox;
    private Config config;

    public MainForm()
    {
        // Initialize the form
        this.Text = "Select Directory to Extract Comments";
        this.Size = new System.Drawing.Size(600, 400);

        selectDirectoryButton = new Button
        {
            Text = "Select Files", //"Select Directory",
            Dock = DockStyle.Top,
            BackColor = System.Drawing.Color.Yellow
        };
        selectDirectoryButton.Click += SelectFiles; //SelectDirectory;

        outputBox = new TextBox { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical };

        Controls.Add(outputBox);
        Controls.Add(selectDirectoryButton);

        // Load configuration
        LoadConfig();
    }

    //private void LoadConfig()
    //{
    //    try
    //    {
    //        string configText = File.ReadAllText("config.json");
    //        config = JsonConvert.DeserializeObject<Config>(configText);
    //    }
    //    catch (Exception ex)
    //    {
    //        MessageBox.Show($"Error loading configuration: {ex.Message}");
    //    }
    //}
    private void LoadConfig()
    {
        try
        {
            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
            string configText = File.ReadAllText(configPath);
            config = JsonConvert.DeserializeObject<Config>(configText);

            // Update the RScriptFilePath to be relative to the config.json directory
            string configDirectory = Path.GetDirectoryName(configPath);
            config.RScriptFilePath = Path.Combine(configDirectory, config.RScriptFilePath);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading configuration: {ex.Message}");
        }
    }

    private void SelectFiles(object sender, EventArgs e)
    {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string[] selectedFiles = openFileDialog.FileNames;
                outputBox.Text = "Selected Files:" + Environment.NewLine;
                foreach (string file in selectedFiles)
                {
                    outputBox.Text += file + Environment.NewLine;
                }
                ProcessDocxFiles(selectedFiles);
            }
        }
    }

    //private void ProcessDocxFiles(string[] filePaths)
    //{
    //    foreach (string filePath in filePaths)
    //    {
    //        string directoryPath = Path.GetDirectoryName(filePath);
    //        RunRScript(filePath, directoryPath);
    //    }
    //}

    private void ProcessDocxFiles(string[] filePaths)
    {
        string directoryPath = Path.GetDirectoryName(filePaths[0]);
        RunRScript(filePaths, directoryPath);
    }

    private void RunRScript(string[] filePaths, string outputDirectory)
    {
        try
        {
            string filePathsArgument = string.Join(" ", filePaths.Select(fp => $"\"{fp}\""));
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = config.RscriptPath,
                //Arguments = $"\"{config.RScriptFilePath}\" \"{filePath}\" \"{outputDirectory}\"",
                Arguments = $"\"{config.RScriptFilePath}\" {filePathsArgument} \"{outputDirectory}\"",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = new Process { StartInfo = psi })
            {
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                string[] outputLines = output.Split(new[] { "!" }, StringSplitOptions.None);
                foreach (string line in outputLines)
                {
                    outputBox.Text += line + Environment.NewLine;
                }
            }
        }
        catch (Exception ex)
        {
            outputBox.Text += $"Error: {ex.Message}" + Environment.NewLine;
        }
    }
}

public static class Program
{
    [STAThread]
    public static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}
