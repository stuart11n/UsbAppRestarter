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

/// <summary>
/// A C# WinForms application that monitors USB device connection events
/// and restarts a specified list of applications upon detection of a new USB device.
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
        this.Text = "USB Device App Restarter";
        this.Width = 550;
        this.Height = 400;
        this.FormBorderStyle = FormBorderStyle.FixedSingle;
        this.MaximizeBox = false;
        this.BackColor = Color.WhiteSmoke;
        this.Font = new Font("Segoe UI", 10);

        // Load settings and sync UI state
        LoadApplicationList();
        CheckAutoStartStatus();
    }

    private void InitializeComponent()
    {
        // Application List Label
        appListLabel = new Label
        {
            Text = "Applications to Restart:",
            Location = new Point(20, 20),
            Width = 490,
            AutoSize = true
        };

        // Application List Box
        appList = new ListBox
        {
            Location = new Point(20, 45),
            Width = 490,
            Height = 150,
            SelectionMode = SelectionMode.One
        };

        // Add Button
        addButton = new Button
        {
            Text = "Add Application",
            Location = new Point(20, 205),
            Width = 150,
            Height = 30,
            BackColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        addButton.Click += AddButton_Click;

        // Remove Button
        removeButton = new Button
        {
            Text = "Remove Selected",
            Location = new Point(180, 205),
            Width = 150,
            Height = 30,
            BackColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        removeButton.Click += RemoveButton_Click;

        // Auto-Start Checkbox
        autoStartCheckbox = new CheckBox
        {
            Text = "Start with Windows Login",
            Location = new Point(350, 210),
            Width = 160,
            Height = 20,
            AutoSize = true
        };
        autoStartCheckbox.CheckedChanged += AutoStartCheckbox_CheckedChanged;

        // Start/Stop Monitor Button 
        startMonitorButton = new Button
        {
            Text = "Start Monitoring",
            Location = new Point(20, 240),
            Width = 490,
            Height = 40,
            BackColor = Color.SeaGreen,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        startMonitorButton.FlatAppearance.BorderSize = 0;
        startMonitorButton.Click += StartMonitorButton_Click;

        // Status Label 
        statusLabel = new Label
        {
            Text = "Status: Not Monitoring",
            Location = new Point(20, 290),
            Width = 490,
            Height = 60,
            BorderStyle = BorderStyle.FixedSingle,
            Padding = new Padding(10),
            TextAlign = ContentAlignment.TopLeft,
            BackColor = Color.LightYellow
        };

        // Add controls to the form
        this.Controls.Add(appListLabel);
        this.Controls.Add(appList);
        this.Controls.Add(addButton);
        this.Controls.Add(removeButton);
        this.Controls.Add(autoStartCheckbox);
        this.Controls.Add(startMonitorButton);
        this.Controls.Add(statusLabel);

        // --- System Tray Initialization ---
        contextMenuStrip = new ContextMenuStrip();
        contextMenuStrip.Items.Add("Show", null, ShowMenuItem_Click);
        contextMenuStrip.Items.Add("-");
        contextMenuStrip.Items.Add("Exit", null, ExitMenuItem_Click);

        notifyIcon = new NotifyIcon();
        notifyIcon.Text = this.Text;
        // Using SystemIcons.Application as a placeholder icon 
        notifyIcon.Icon = SystemIcons.Application;
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


    // --- List Management ---
    private void LoadApplicationList()
    {
        appList.Items.Clear();
        var paths = Properties.Settings.Default.GetExecutablePaths();
        foreach (string path in paths)
        {
            appList.Items.Add(path);
        }
    }

    private void SaveApplicationList()
    {
        List<string> paths = appList.Items.Cast<string>().ToList();
        Properties.Settings.Default.SetExecutablePaths(paths);
        Properties.Settings.Default.Save();
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
        if (appList.Items.Count == 0)
        {
            SetStatus("Error: Please add at least one application path to the list.", Color.Salmon);
            return;
        }

        // Validate all paths before starting
        List<string> validPaths = new List<string>();
        foreach (string path in appList.Items)
        {
            if (File.Exists(path))
            {
                validPaths.Add(path);
            }
            else
            {
                // Note: SetStatus in this loop may cause slow startup if many paths are invalid.
                // We proceed and show a general warning.
            }
        }

        if (validPaths.Count == 0)
        {
            SetStatus("Error: No valid executable paths found in the list.", Color.Salmon);
            // Since monitoring didn't start, re-enable UI elements
            EnableConfigurationUI(true);
            return;
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

            SetStatus($"Monitoring started for {validPaths.Count} application(s). Waiting for USB connection...", Color.LightGreen);
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
                deviceName = instance["Caption"]?.ToString() ?? "Generic PnP Device";

                SetStatus($"USB Device Connected: '{deviceName}'. Attempting to restart {appList.Items.Count} application(s)...", Color.LightBlue);

                // Get the list of applications to restart
                List<string> pathsToRestart = appList.Items.Cast<string>().ToList();

                // Call the restart logic for the whole list
                RestartApplications(pathsToRestart);
            }
            catch (Exception ex)
            {
                SetStatus($"Error processing device event: {ex.Message}", Color.Salmon);
            }
        }));
    }

    // --- Application Restart Logic ---
    private void RestartApplications(List<string> executablePaths)
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
            SetStatus($"Successfully restarted {successfulRestarts} application(s) due to USB connection.", Color.LightGreen);
        }
    }

    // --- Auto Start Logic ---
    private void CheckAutoStartStatus()
    {
        // Read setting from the custom settings class (loaded from JSON)
        bool savedState = Properties.Settings.Default.GetAutoStartEnabled();
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
                Properties.Settings.Default.SetAutoStartEnabled(true);
                Properties.Settings.Default.Save();
                SetStatus("Auto-Start enabled. App will run on Windows login.", Color.LightGreen);
            }
            else
            {
                // If setting the registry fails (e.g., permissions), revert checkbox state
                autoStartCheckbox.Checked = false;
                Properties.Settings.Default.SetAutoStartEnabled(false);
                Properties.Settings.Default.Save();
                SetStatus("Error enabling Auto-Start. Check permissions.", Color.Salmon);
            }
        }
        else
        {
            SetRegistryRunKey(false);
            Properties.Settings.Default.SetAutoStartEnabled(false);
            Properties.Settings.Default.Save();
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

        // Clean up NotifyIcon resources
        if (notifyIcon != null)
        {
            notifyIcon.Visible = false;
            notifyIcon.Dispose();
        }
    }
}

// --- Custom Settings Implementation ---
namespace Properties
{
    internal class Settings
    {
        // Define the JSON file path for persistence in the application directory
        private static readonly string SettingsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app_restart_paths.json");

        // Internal storage for the serialized string and the new setting
        private string _serializedPaths = "";
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
                // Write both settings into a simple JSON structure to the file
                string escapedPaths = _serializedPaths.Replace("\\", "\\\\"); // Escape backslashes for JSON compatibility

                // JSON Structure: {"paths": "path1|path2", "autoStart": true/false}
                string jsonContent = "{\"paths\": \"" + escapedPaths + "\", \"autoStart\": " + _autoStartEnabled.ToString().ToLower() + "}";

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

                    // 1. Simple parsing for Paths
                    string searchKeyPaths = "\"paths\": \"";
                    int startPaths = jsonContent.IndexOf(searchKeyPaths) + searchKeyPaths.Length;
                    int endPaths = jsonContent.IndexOf("\"", startPaths);

                    if (startPaths > searchKeyPaths.Length - 1 && endPaths > startPaths)
                    {
                        _serializedPaths = jsonContent.Substring(startPaths, endPaths - startPaths).Replace("\\\\", "\\"); // Unescape backslashes
                    }
                    else
                    {
                        _serializedPaths = "";
                    }

                    // 2. Simple parsing for AutoStart
                    string searchKeyAutoStart = "\"autoStart\": ";
                    int startAutoStart = jsonContent.IndexOf(searchKeyAutoStart) + searchKeyAutoStart.Length;

                    // Find the end of the boolean value (which should be just before the closing brace '}')
                    int endAutoStart = jsonContent.IndexOf("}", startAutoStart);

                    if (startAutoStart > searchKeyAutoStart.Length - 1 && endAutoStart > startAutoStart)
                    {
                        string boolString = jsonContent.Substring(startAutoStart, endAutoStart - startAutoStart).Trim().Replace("\"", "");
                        _autoStartEnabled = bool.TryParse(boolString, out bool result) ? result : false;
                    }
                    else
                    {
                        _autoStartEnabled = false;
                    }

                    Console.WriteLine($"Settings loaded from: {SettingsFilePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading settings from file: {ex.Message}");
                    _serializedPaths = "";
                    _autoStartEnabled = false;
                }
            }
            else
            {
                _serializedPaths = "";
                _autoStartEnabled = false;
            }
        }

        public List<string> GetExecutablePaths()
        {
            // Split the internal string into paths, filtering out empty entries
            return _serializedPaths.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries)
                                  .ToList();
        }

        public void SetExecutablePaths(List<string> paths)
        {
            // Join the list of paths into a single string with a separator
            _serializedPaths = string.Join("|", paths);
        }

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
