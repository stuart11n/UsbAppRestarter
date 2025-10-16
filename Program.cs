using System;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using System.Management;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Text.RegularExpressions; // Required for regex matching

/// <summary>
/// A C# WinForms application that monitors USB device connection events
/// and restarts a specified list of applications upon detection of a new USB device.
/// Restarts are only triggered if the device's name matches one of the user-defined regex filters.
/// It starts minimized to the system tray with monitoring enabled automatically.
/// </summary>
public class UsbRestartMonitorForm : Form
{
    private Label appListLabel;
    private ListBox appList;
    private Button addButton;
    private Button removeButton;
    private CheckBox autoStartCheckbox;
    private Button startMonitorButton;
    private Label statusLabel;

    // New controls for regex filtering
    private Label regexLabel;
    private TextBox regexFilterTextBox;

    private ManagementEventWatcher usbWatcher;
    private bool isMonitoring = false;

    private NotifyIcon notifyIcon;
    private ContextMenuStrip contextMenuStrip;

    // --- Application Entry Point ---
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        // Ensure the System.Management DLL and System.Threading.Tasks are accessible.
        Application.Run(new UsbRestartMonitorForm());
    }

    // --- Constructor and UI Setup ---
    public UsbRestartMonitorForm()
    {
        InitializeComponent();
        // Application title
        this.Text = "USB Device App Restarter";
        // Increased height for new control
        this.Width = 550;
        this.Height = 550;
        this.FormBorderStyle = FormBorderStyle.FixedSingle;
        this.MaximizeBox = false;
        this.BackColor = Color.WhiteSmoke;
        this.Font = new Font("Segoe UI", 10);

        // Load settings and sync UI state
        LoadApplicationList();
        LoadRegexFilters(); // Load the stored regex patterns
        CheckAutoStartStatus();
    }

    private void InitializeComponent()
    {
        using (var stream = new MemoryStream(AppRestarter.Properties.Resources.icon))
        {
            this.Icon = new Icon(stream);
        }

        // --- 1. Application List Group ---
        appListLabel = new Label
        {
            Text = "Applications to Restart:",
            Location = new Point(20, 20),
            Width = 490,
            AutoSize = true
        };

        appList = new ListBox
        {
            Location = new Point(20, 45),
            Width = 490,
            Height = 120, // Shorter list box
            SelectionMode = SelectionMode.One
        };

        addButton = new Button
        {
            Text = "Add Application",
            Location = new Point(20, 175),
            Width = 150,
            Height = 30,
            BackColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        addButton.Click += AddButton_Click;

        removeButton = new Button
        {
            Text = "Remove Selected",
            Location = new Point(180, 175),
            Width = 150,
            Height = 30,
            BackColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        removeButton.Click += RemoveButton_Click;

        autoStartCheckbox = new CheckBox
        {
            Text = "Start with Windows Login",
            Location = new Point(350, 180),
            Width = 160,
            Height = 20,
            AutoSize = true
        };
        autoStartCheckbox.CheckedChanged += AutoStartCheckbox_CheckedChanged;

        // --- 2. Regex Filter Group (NEW) ---
        regexLabel = new Label
        {
            Text = "USB Device ID Filters (Regex - One per line):",
            Location = new Point(20, 220),
            Width = 490,
            AutoSize = true
        };

        regexFilterTextBox = new TextBox
        {
            Location = new Point(20, 245),
            Width = 490,
            Height = 100,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };

        // --- 3. Monitor and Status Group ---
        startMonitorButton = new Button
        {
            Text = "Start Monitoring",
            Location = new Point(20, 355), // Moved down
            Width = 490,
            Height = 40,
            BackColor = Color.SeaGreen,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        startMonitorButton.FlatAppearance.BorderSize = 0;
        startMonitorButton.Click += StartMonitorButton_Click;

        statusLabel = new Label
        {
            Text = "Status: Not Monitoring",
            Location = new Point(20, 410), // Moved down
            Width = 490,
            Height = 60,
            BorderStyle = BorderStyle.FixedSingle,
            Padding = new Padding(10),
            TextAlign = ContentAlignment.TopLeft,
            BackColor = Color.LightYellow
        };

        // Add all controls to the form
        this.Controls.Add(appListLabel);
        this.Controls.Add(appList);
        this.Controls.Add(addButton);
        this.Controls.Add(removeButton);
        this.Controls.Add(autoStartCheckbox);
        this.Controls.Add(regexLabel); // Add new label
        this.Controls.Add(regexFilterTextBox); // Add new textbox
        this.Controls.Add(startMonitorButton);
        this.Controls.Add(statusLabel);

        // --- System Tray Initialization ---
        contextMenuStrip = new ContextMenuStrip();
        contextMenuStrip.Items.Add("Show", null, ShowMenuItem_Click);
        contextMenuStrip.Items.Add("-");
        contextMenuStrip.Items.Add("Exit", null, ExitMenuItem_Click);

        notifyIcon = new NotifyIcon();
        notifyIcon.Text = "USB Device App Restarter";
        using (var stream = new MemoryStream(AppRestarter.Properties.Resources.icon))
        {
            notifyIcon.Icon = new Icon(stream);
        }

        notifyIcon.ContextMenuStrip = contextMenuStrip;
        notifyIcon.Visible = false;

        // Handler for double-click restore
        notifyIcon.MouseDoubleClick += NotifyIcon_MouseDoubleClick;

        // Hide from taskbar initially
        this.ShowInTaskbar = false;
    }

    // --- Form Overrides for Tray and Auto-Start ---
    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);

        // 1. Hide the window and show the tray icon
        this.Hide();
        if (notifyIcon != null)
        {
            notifyIcon.Visible = true;
            // Show a helpful balloon tip on startup
            notifyIcon.ShowBalloonTip(1000, this.Text, "Application is running in the background. Double-click to restore.", ToolTipIcon.Info);
        }

        // 2. Start monitoring automatically when the application launches
        StartMonitoring();
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        // Intercept minimize action to hide to tray
        if (this.WindowState == FormWindowState.Minimized)
        {
            this.Hide();
            if (notifyIcon != null)
            {
                notifyIcon.Visible = true;
            }
        }
    }

    // --- Tray Icon Handlers ---
    private void ShowMenuItem_Click(object sender, EventArgs e)
    {
        RestoreForm();
    }

    private void NotifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
    {
        RestoreForm();
    }

    private void RestoreForm()
    {
        this.Show();
        this.WindowState = FormWindowState.Normal;
        this.ShowInTaskbar = true;
        if (notifyIcon != null)
        {
            notifyIcon.Visible = false;
        }
    }

    private void ExitMenuItem_Click(object sender, EventArgs e)
    {
        // This ensures the application shuts down completely
        Application.Exit();
    }


    // --- List and Regex Management ---
    private void LoadApplicationList()
    {
        appList.Items.Clear();
        var paths = InternalProperties.Settings.Default.GetExecutablePaths();
        foreach (string path in paths)
        {
            appList.Items.Add(path);
        }
    }

    private void LoadRegexFilters()
    {
        string patterns = InternalProperties.Settings.Default.GetRegexFilterString();
        regexFilterTextBox.Text = patterns;
    }

    private void SaveApplicationList()
    {
        List<string> paths = appList.Items.Cast<string>().ToList();
        InternalProperties.Settings.Default.SetExecutablePaths(paths);
        InternalProperties.Settings.Default.Save();
    }

    private void SaveRegexFilters()
    {
        InternalProperties.Settings.Default.SetRegexFilterString(regexFilterTextBox.Text);
        InternalProperties.Settings.Default.Save();
    }

    private void AddButton_Click(object sender, EventArgs e)
    {
        using (OpenFileDialog ofd = new OpenFileDialog())
        {
            ofd.Filter = "Executable Files (*.exe)|*.exe";
            ofd.Title = "Select the application to restart";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                if (!appList.Items.Contains(ofd.FileName))
                {
                    appList.Items.Add(ofd.FileName);
                    SaveApplicationList();
                    SetStatus($"Added: {Path.GetFileName(ofd.FileName)}.", Color.WhiteSmoke);
                }
                else
                {
                    SetStatus("Application is already in the list.", Color.LightCoral);
                }
            }
        }
    }

    private void RemoveButton_Click(object sender, EventArgs e)
    {
        if (appList.SelectedItem != null)
        {
            string removedItem = appList.SelectedItem.ToString();
            appList.Items.Remove(appList.SelectedItem);
            SaveApplicationList();
            SetStatus($"Removed: {Path.GetFileName(removedItem)}.", Color.WhiteSmoke);
        }
        else
        {
            SetStatus("Please select an application to remove.", Color.LightCoral);
        }
    }

    // --- Monitoring Control ---
    private void StartMonitorButton_Click(object sender, EventArgs e)
    {
        if (isMonitoring)
        {
            StopMonitoring();
        }
        else
        {
            StartMonitoring();
        }
    }

    private void StartMonitoring()
    {
        // Save current regex filters and app paths before starting the monitor
        SaveApplicationList();
        SaveRegexFilters();

        List<string> validAppPaths = InternalProperties.Settings.Default.GetExecutablePaths();

        if (validAppPaths.Count == 0)
        {
            SetStatus("Error: Please add at least one valid application path to the list.", Color.Salmon);
            return;
        }

        // 1. Validate application paths
        List<string> existingAppPaths = validAppPaths.Where(p => File.Exists(p)).ToList();

        if (existingAppPaths.Count == 0)
        {
            SetStatus("Error: No valid executable paths found on disk.", Color.Salmon);
            EnableConfigurationUI(true);
            return;
        }

        // 2. Validate regex filters (optional check, but good practice)
        List<string> filters = InternalProperties.Settings.Default.GetRegexFilters();
        if (filters.Count == 0)
        {
            SetStatus("Warning: No regex filters defined. Monitoring ALL USB devices. Add filters for specific devices.", Color.Yellow);
        }

        try
        {
            // WMI Query Language (WQL) to detect a new device (like USB) insertion
            string wmiQuery = "SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_PnPEntity'";

            usbWatcher = new ManagementEventWatcher(wmiQuery);
            usbWatcher.EventArrived += UsbDevice_EventArrived;
            usbWatcher.Start();

            isMonitoring = true;

            // Update UI elements for running state
            startMonitorButton.Text = "Stop Monitoring";
            startMonitorButton.BackColor = Color.DarkRed;
            EnableConfigurationUI(false);

            SetStatus($"Monitoring started for {existingAppPaths.Count} application(s). Filters: {filters.Count} active.", Color.LightGreen);
        }
        catch (Exception ex)
        {
            SetStatus($"Failed to start WMI monitoring: {ex.Message}", Color.Salmon);
            usbWatcher?.Stop();
            usbWatcher = null;
            EnableConfigurationUI(true);
        }
    }

    private void EnableConfigurationUI(bool enable)
    {
        addButton.Enabled = enable;
        removeButton.Enabled = enable;
        appList.Enabled = enable;
        autoStartCheckbox.Enabled = enable;
        regexFilterTextBox.Enabled = enable; // Disable regex input while monitoring
    }


    private void StopMonitoring()
    {
        usbWatcher?.Stop();
        usbWatcher = null;

        isMonitoring = false;

        // Update UI elements for stopped state
        startMonitorButton.Text = "Start Monitoring";
        startMonitorButton.BackColor = Color.SeaGreen;
        EnableConfigurationUI(true);

        SetStatus("Status: Monitoring stopped.", Color.LightYellow);
    }

    // --- WMI Event Handler ---
    private void UsbDevice_EventArrived(object sender, EventArrivedEventArgs e)
    {
        // Run the restart logic on the main UI thread via Invoke
        this.Invoke(new Action(() =>
        {
            string deviceName = "Unknown Device";
            try
            {
                ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
                deviceName = instance["DeviceId"]?.ToString();

                List<string> filters = InternalProperties.Settings.Default.GetRegexFilters();
                SetStatus($"Invalid regex pattern detected: '{deviceName}'. Please correct it.", Color.Salmon);
            
                // 1. Apply Regex Filter Logic
                if (filters.Any())
                {
                    bool matchFound = false;
                    foreach (string pattern in filters)
                    {
                        if (string.IsNullOrWhiteSpace(pattern)) continue;

                        try
                        {
                            // Check if the device name matches the pattern (case-insensitive)
                            if (Regex.IsMatch(deviceName, pattern, RegexOptions.IgnoreCase))
                            {
                                matchFound = true;
                                break;
                            }
                        }
                        catch (ArgumentException)
                        {
                            // Invalid regex pattern. Log this but continue trying other filters.
                            SetStatus($"Invalid regex pattern detected: '{pattern}'. Please correct it.", Color.Salmon);
                        }
                    }

                    if (!matchFound)
                    {
                        SetStatus($"USB Connected: '{deviceName}'. No matching regex filter found. Skipping restart.", Color.Orange);
                        return; // EXIT if no match
                    }
                }

                // 2. Proceed with restart if no filters defined or if a match was found
                SetStatus($"USB Connected: '{deviceName}' (Match found). Attempting to restart {appList.Items.Count} application(s)...", Color.LightBlue);

                List<string> pathsToRestart = InternalProperties.Settings.Default.GetExecutablePaths();
                RestartApplications(pathsToRestart, deviceName);
            }
            catch (Exception ex)
            {
                SetStatus($"Error processing device event: {ex.Message}", Color.Salmon);
            }
        }));
    }

    // --- Application Restart Logic ---
    private void RestartApplications(List<string> executablePaths, string dev)
    {
        int successfulRestarts = 0;
        foreach (string executablePath in executablePaths)
        {
            if (!File.Exists(executablePath))
            {
                SetStatus($"Skipping restart: File not found at '{Path.GetFileName(executablePath)}'.", Color.LightCoral);
                continue;
            }

            try
            {
                string processName = Path.GetFileNameWithoutExtension(executablePath);
                Process[] existingProcesses = Process.GetProcessesByName(processName);

                // 1. Terminate existing process (if running)
                if (existingProcesses.Length > 0)
                {
                    foreach (Process p in existingProcesses)
                    {
                        if (!p.HasExited) p.Kill();
                    }
                    // Wait briefly for processes to exit
                    System.Threading.Thread.Sleep(200);
                }

                // 2. Start the application
                Process.Start(executablePath);
                successfulRestarts++;
            }
            catch (Exception ex)
            {
                SetStatus($"Error restarting '{Path.GetFileName(executablePath)}': {ex.Message}", Color.Salmon);
            }
        }

        if (successfulRestarts > 0)
        {
            SetStatus($"Successfully restarted {successfulRestarts} application(s) due to {dev}.", Color.LightGreen);
        }
    }

    // --- Auto Start Logic ---
    private void CheckAutoStartStatus()
    {
        // Read setting from the custom settings class (loaded from JSON)
        bool savedState = InternalProperties.Settings.Default.GetAutoStartEnabled();
        autoStartCheckbox.Checked = savedState;

        // Ensure the registry reflects the saved state on startup
        if (savedState)
        {
            SetRegistryRunKey(true);
        }
    }

    private void AutoStartCheckbox_CheckedChanged(object sender, EventArgs e)
    {
        if (autoStartCheckbox.Checked)
        {
            if (SetRegistryRunKey(true))
            {
                InternalProperties.Settings.Default.SetAutoStartEnabled(true);
                InternalProperties.Settings.Default.Save();
                SetStatus("Auto-Start enabled. App will run on Windows login.", Color.LightGreen);
            }
            else
            {
                // If setting the registry fails (e.g., permissions), revert checkbox state
                autoStartCheckbox.Checked = false;
                InternalProperties.Settings.Default.SetAutoStartEnabled(false);
                InternalProperties.Settings.Default.Save();
                SetStatus("Error enabling Auto-Start. Check permissions.", Color.Salmon);
            }
        }
        else
        {
            SetRegistryRunKey(false);
            InternalProperties.Settings.Default.SetAutoStartEnabled(false);
            InternalProperties.Settings.Default.Save();
            SetStatus("Auto-Start disabled.", Color.LightYellow);
        }
    }

    private bool SetRegistryRunKey(bool enable)
    {
        const string runKeyName = "UsbRestartMonitor";
        // Use HKEY_CURRENT_USER for permissions simplicity
        RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true);

        if (key == null) return false;

        try
        {
            if (enable)
            {
                // Set the value to the current executable path
                string appPath = Application.ExecutablePath;
                key.SetValue(runKeyName, appPath);
                return true;
            }
            else
            {
                // Remove the value
                if (key.GetValue(runKeyName) != null)
                {
                    key.DeleteValue(runKeyName);
                }
                return true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Registry error: {ex.Message}");
            return false;
        }
        finally
        {
            key.Close();
        }
    }

    // --- Utility Methods ---
    private void SetStatus(string message, Color backgroundColor)
    {
        // Use BeginInvoke for non-blocking UI update if not on the main thread
        if (this.InvokeRequired)
        {
            this.BeginInvoke(new Action(() => SetStatus(message, backgroundColor)));
            return;
        }

        statusLabel.Text = $"Status: {message}";
        statusLabel.BackColor = backgroundColor;
        // Optionally flash the background color for a moment
        Task.Delay(1000).ContinueWith(_ =>
        {
            // Only revert color if it hasn't changed to a new status color
            if (statusLabel.BackColor == backgroundColor)
            {
                statusLabel.BackColor = Color.WhiteSmoke;
            }
        }, TaskScheduler.FromCurrentSynchronizationContext());
    }

    // Clean up WMI watcher and NotifyIcon on form close
    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        base.OnFormClosing(e);
        StopMonitoring();

        // Save settings on close to ensure latest list/regex are preserved
        SaveApplicationList();
        SaveRegexFilters();

        // Clean up NotifyIcon resources
        if (notifyIcon != null)
        {
            notifyIcon.Visible = false;
            notifyIcon.Dispose();
        }
    }
}

// --- Custom Settings Implementation ---
namespace InternalProperties
{
    internal class Settings
    {
        // Define the JSON file path for persistence in the application directory
        private static readonly string SettingsFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "USB Device App Restarter","app_restart_paths.json");
        // Internal storage for the serialized settings
        private string _serializedPaths = "";
        private string _regexFilterString = ""; // New internal storage for regex filters
        private bool _autoStartEnabled = false;

        public static Settings Default { get; } = new Settings();

        // Constructor now handles loading from the file
        private Settings()
        {
            LoadFromFile();
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsFilePath));

                // Escape backslashes in paths and filters for JSON compatibility
                string escapedPaths = _serializedPaths.Replace("\\", "\\\\");
                // Escape characters in the regex string itself
                string escapedFilters = _regexFilterString.Replace("\\", "\\\\").Replace("\"", "\\\"");

                // JSON Structure: {"paths": "path1|path2", "autoStart": true/false, "regexFilters": "pattern1\npattern2"}
                string jsonContent = "{"
                    + $"\"paths\": \"{escapedPaths}\", "
                    + $"\"autoStart\": {(_autoStartEnabled ? "true" : "false")}, "
                    + $"\"regexFilters\": \"{escapedFilters}\""
                    + "}";

                File.WriteAllText(SettingsFilePath, jsonContent);
                Console.WriteLine($"Settings saved to: {SettingsFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving settings to file: {ex.Message}");
            }
        }

        private void LoadFromFile()
        {
            if (File.Exists(SettingsFilePath))
            {
                try
                {
                    string jsonContent = File.ReadAllText(SettingsFilePath);

                    // Simple, robust parsing logic (using string manipulation)
                    Func<string, string> extractValue = (key) =>
                    {
                        string searchKey = $"\"{key}\": \"";
                        int start = jsonContent.IndexOf(searchKey) + searchKey.Length;
                        if (start >= searchKey.Length)
                        {
                            int end = jsonContent.IndexOf("\"", start);
                            if (end > start)
                            {
                                return jsonContent.Substring(start, end - start).Replace("\\\\", "\\");
                            }
                        }
                        return "";
                    };

                    // 1. Load Paths
                    _serializedPaths = extractValue("paths");

                    // 2. Load Regex Filters
                    _regexFilterString = extractValue("regexFilters");

                    // 3. Load AutoStart (boolean parsing)
                    string autoStartKey = "\"autoStart\": ";
                    int startAutoStart = jsonContent.IndexOf(autoStartKey) + autoStartKey.Length;
                    if (startAutoStart >= autoStartKey.Length)
                    {
                        int endAutoStart = jsonContent.IndexOf(",", startAutoStart);
                        if (endAutoStart == -1) endAutoStart = jsonContent.IndexOf("}", startAutoStart); // Handle if it's the last element

                        if (endAutoStart > startAutoStart)
                        {
                            string boolString = jsonContent.Substring(startAutoStart, endAutoStart - startAutoStart).Trim();
                            _autoStartEnabled = bool.TryParse(boolString, out bool result) && result;
                        }
                        else
                        {
                            _autoStartEnabled = false;
                        }
                    }

                    Console.WriteLine($"Settings loaded from: {SettingsFilePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading settings from file: {ex.Message}");
                    _serializedPaths = "";
                    _regexFilterString = "";
                    _autoStartEnabled = false;
                }
            }
            else
            {
                _serializedPaths = "";
                _regexFilterString = "";
                _autoStartEnabled = false;
            }
        }

        // --- Paths Accessors ---
        public List<string> GetExecutablePaths()
        {
            return _serializedPaths.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries)
                                  .ToList();
        }

        public void SetExecutablePaths(List<string> paths)
        {
            _serializedPaths = string.Join("|", paths);
        }

        // --- Regex Accessors (NEW) ---
        // Returns the raw string content for the TextBox
        public string GetRegexFilterString()
        {
            return _regexFilterString;
        }

        // Sets the raw string content from the TextBox
        public void SetRegexFilterString(string patterns)
        {
            _regexFilterString = patterns;
        }

        // Returns a parsed list of patterns (one per line)
        public List<string> GetRegexFilters()
        {
            // Split by newline and filter out any empty lines/patterns
            return _regexFilterString.Split(new char[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                     .Select(s => s.Trim())
                                     .Where(s => !string.IsNullOrEmpty(s))
                                     .ToList();
        }

        // --- Auto Start Accessors ---
        public bool GetAutoStartEnabled()
        {
            return _autoStartEnabled;
        }

        public void SetAutoStartEnabled(bool enabled)
        {
            _autoStartEnabled = enabled;
        }
    }
}
