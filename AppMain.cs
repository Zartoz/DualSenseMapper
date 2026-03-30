using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using HidLibrary;
using Nefarius.ViGEm.Client;
using Nefarius.ViGEm.Client.Exceptions;
using Nefarius.ViGEm.Client.Targets;
using Nefarius.ViGEm.Client.Targets.Xbox360;
using System.Windows.Forms;

const int SonyVendorId = 0x054C;
const int ReadTimeoutMs = 1000;
const uint RequestedTimerResolutionMs = 1;
const short GenericDesktopUsagePage = 0x01;
const short JoystickUsage = 0x04;
const short GamepadUsage = 0x05;
const int PreferredHidInputBufferCount = 3;
const int DefaultLightbarHue = 220;
const byte DualSenseOutputReportUsbId = 0x02;
const int DualSenseOutputReportUsbLength = 63;
const byte DualSenseOutputReportBluetoothId = 0x31;
const int DualSenseOutputReportBluetoothLength = 78;
const byte DualSenseOutputTag = 0x10;
const byte DualSenseLightbarControlEnableFlag = 0x04;
const byte DualSenseLightbarSetupControlEnableFlag = 0x02;
const byte DualSenseLightbarSetupLightOut = 0x02;
const byte DualSenseOutputCrcSeed = 0xA2;
const int DualSenseUsbValidFlag1Byte = 2;
const int DualSenseUsbValidFlag2Byte = 39;
const int DualSenseUsbLightbarSetupByte = 42;
const int DualSenseUsbLightbarRedByte = 45;
const int DualSenseUsbLightbarGreenByte = 46;
const int DualSenseUsbLightbarBlueByte = 47;
const int DualSenseBluetoothValidFlag1Byte = 4;
const int DualSenseBluetoothValidFlag2Byte = 41;
const int DualSenseBluetoothLightbarSetupByte = 44;
const int DualSenseBluetoothLightbarRedByte = 47;
const int DualSenseBluetoothLightbarGreenByte = 48;
const int DualSenseBluetoothLightbarBlueByte = 49;

const int LeftStickXByte = 1;
const int LeftStickYByte = 2;
const int RightStickXByte = 3;
const int RightStickYByte = 4;
const int LeftTriggerByte = 5;
const int RightTriggerByte = 6;
const int Buttons1Byte = 8;
const int Buttons2Byte = 9;
const int Buttons3Byte = 10;
const byte DPadMask = 0x0F;
const byte SquareMask = 0x10;
const byte CrossMask = 0x20;
const byte CircleMask = 0x40;
const byte TriangleMask = 0x80;
const byte LeftShoulderMask = 0x01;
const byte RightShoulderMask = 0x02;
const byte LeftThumbMask = 0x40;
const byte RightThumbMask = 0x80;
const byte OptionsMask = 0x20;
const byte ShareMask = 0x10;
const byte PsMask = 0x01;
const byte TouchpadMask = 0x02;
const bool InvertLeftStickY = true;
const bool InvertRightStickY = true;

var devices = new List<(HidDevice Device, string Product, string Manufacturer, int InputLength, int OutputLength, int FeatureLength)>();
CancellationTokenSource? activeRunCts = null;
Task? activeRunTask = null;
var currentMode = "idle";
var lightbarHue = DefaultLightbarHue;
var lightbarColorArgb = ColorFromHue(lightbarHue).ToArgb();
var isClosingAfterStop = false;
byte dualSenseOutputSequence = 0;

ApplicationConfiguration.Initialize();

var form = new Form
{
    Text = "DualSense Mapper",
    StartPosition = FormStartPosition.CenterScreen,
    Width = 760,
    Height = 520,
    MinimumSize = new Size(680, 460),
    Font = new Font("Segoe UI", 10),
    BackColor = Color.FromArgb(245, 247, 250)
};

var root = new TableLayoutPanel
{
    Dock = DockStyle.Fill,
    ColumnCount = 1,
    RowCount = 7,
    Padding = new Padding(16),
    BackColor = form.BackColor
};
root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
form.Controls.Add(root);

var titleLabel = new Label
{
    AutoSize = true,
    Font = new Font("Segoe UI Semibold", 18, FontStyle.Bold),
    Text = "DualSense to Xbox 360 Mapper",
    Margin = new Padding(0, 0, 0, 4)
};

var subtitleLabel = new Label
{
    AutoSize = true,
    MaximumSize = new Size(700, 0),
    ForeColor = Color.FromArgb(85, 92, 102),
    Text = "Minimal build. Supports DualSense over USB and Bluetooth with low-latency HID settings.",
    Margin = new Padding(0, 0, 0, 12)
};

var deviceGroup = new GroupBox
{
    Dock = DockStyle.Top,
    AutoSize = true,
    Text = "Controller"
};

var deviceLayout = new TableLayoutPanel
{
    Dock = DockStyle.Top,
    AutoSize = true,
    ColumnCount = 2,
    RowCount = 2,
    Padding = new Padding(10)
};
deviceLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
deviceLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
deviceLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
deviceLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
deviceGroup.Controls.Add(deviceLayout);

var deviceCombo = new ComboBox
{
    Dock = DockStyle.Top,
    DropDownStyle = ComboBoxStyle.DropDownList,
    Width = 520,
    Margin = new Padding(0, 0, 8, 8)
};

var refreshButton = new Button
{
    AutoSize = true,
    Text = "Refresh",
    Padding = new Padding(10, 4, 10, 4),
    Margin = new Padding(0, 0, 0, 8)
};

var deviceDetailsLabel = new Label
{
    AutoSize = true,
    MaximumSize = new Size(660, 0),
    ForeColor = Color.FromArgb(75, 82, 92),
    Text = "No device selected."
};

deviceLayout.Controls.Add(deviceCombo, 0, 0);
deviceLayout.Controls.Add(refreshButton, 1, 0);
deviceLayout.Controls.Add(deviceDetailsLabel, 0, 1);
deviceLayout.SetColumnSpan(deviceDetailsLabel, 2);

var lightbarGroup = new GroupBox
{
    Dock = DockStyle.Top,
    AutoSize = true,
    Text = "DualSense LED"
};

var lightbarLayout = new TableLayoutPanel
{
    Dock = DockStyle.Top,
    AutoSize = true,
    ColumnCount = 3,
    RowCount = 2,
    Padding = new Padding(10)
};
lightbarLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
lightbarLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
lightbarLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
lightbarLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
lightbarLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
lightbarGroup.Controls.Add(lightbarLayout);

var lightbarLabel = new Label
{
    AutoSize = true,
    Text = "Hue",
    Anchor = AnchorStyles.Left,
    Margin = new Padding(0, 8, 10, 0)
};

var lightbarSlider = new TrackBar
{
    Dock = DockStyle.Top,
    Minimum = 0,
    Maximum = 359,
    TickFrequency = 30,
    LargeChange = 15,
    SmallChange = 1,
    Value = lightbarHue,
    Margin = new Padding(0, 0, 10, 0)
};

var lightbarPreview = new Panel
{
    Width = 44,
    Height = 24,
    Margin = new Padding(0, 8, 0, 0),
    BorderStyle = BorderStyle.FixedSingle
};

var lightbarValueLabel = new Label
{
    AutoSize = true,
    ForeColor = Color.FromArgb(75, 82, 92),
    Margin = new Padding(0, 4, 0, 0)
};

lightbarLayout.Controls.Add(lightbarLabel, 0, 0);
lightbarLayout.Controls.Add(lightbarSlider, 1, 0);
lightbarLayout.Controls.Add(lightbarPreview, 2, 0);
lightbarLayout.Controls.Add(lightbarValueLabel, 1, 1);
lightbarLayout.SetColumnSpan(lightbarValueLabel, 2);

var actionPanel = new FlowLayoutPanel
{
    Dock = DockStyle.Top,
    AutoSize = true,
    WrapContents = false,
    Margin = new Padding(0, 12, 0, 12)
};

var mapButton = new Button
{
    AutoSize = true,
    Text = "Start Mapping",
    Padding = new Padding(12, 5, 12, 5),
    Margin = new Padding(0, 0, 8, 0)
};

var stopButton = new Button
{
    AutoSize = true,
    Text = "Stop",
    Padding = new Padding(12, 5, 12, 5),
    Enabled = false
};

actionPanel.Controls.Add(mapButton);
actionPanel.Controls.Add(stopButton);

var statusLabel = new Label
{
    AutoSize = true,
    Padding = new Padding(10, 6, 10, 6),
    BackColor = Color.FromArgb(235, 240, 247),
    ForeColor = Color.FromArgb(43, 63, 96),
    Text = "Status: Idle",
    Margin = new Padding(0, 0, 0, 12)
};

var logBox = new TextBox
{
    Dock = DockStyle.Fill,
    Multiline = true,
    ReadOnly = true,
    ScrollBars = ScrollBars.Vertical,
    BackColor = Color.White,
    Font = new Font("Consolas", 9.5f),
    BorderStyle = BorderStyle.FixedSingle
};

root.Controls.Add(titleLabel, 0, 0);
root.Controls.Add(subtitleLabel, 0, 1);
root.Controls.Add(deviceGroup, 0, 2);
root.Controls.Add(lightbarGroup, 0, 3);
root.Controls.Add(actionPanel, 0, 4);
var logHost = new Panel
{
    Dock = DockStyle.Fill,
    Margin = new Padding(0)
};
logHost.Controls.Add(logBox);
root.Controls.Add(statusLabel, 0, 5);
root.Controls.Add(logHost, 0, 6);

refreshButton.Click += (_, _) => RefreshDevices();
deviceCombo.SelectedIndexChanged += (_, _) => UpdateSelectedDeviceDetails();
lightbarSlider.ValueChanged += (_, _) => HandleLightbarSliderChanged();

mapButton.Click += async (_, _) =>
{
    if (currentMode == "mapping")
    {
        return;
    }

    var selected = GetSelectedDevice();
    if (selected is null)
    {
        MessageBox.Show(form, "Connect your DualSense by USB or Bluetooth and click Refresh first.", "No device", MessageBoxButtons.OK, MessageBoxIcon.Information);
        return;
    }

    await StopCurrentRunAsync();
    AppendLog($"Starting mapping on PID 0x{selected.Value.Device.Attributes.ProductId:X4} over {GetConnectionLabel(selected.Value.Device)}.");
    StartRun("mapping", "Mapping running", token => RunMappedMode(selected.Value.Device, token));
};

stopButton.Click += async (_, _) => await StopCurrentRunAsync();
form.FormClosing += async (_, e) =>
{
    if (isClosingAfterStop || activeRunCts is null)
    {
        return;
    }

    e.Cancel = true;
    isClosingAfterStop = true;
    SetStatus("Stopping...");
    await StopCurrentRunAsync();

    if (!form.IsDisposed)
    {
        form.BeginInvoke(new Action(form.Close));
    }
};

UpdateLightbarPreview();
RefreshDevices();
Application.Run(form);

void RefreshDevices()
{
    devices = HidDevices.Enumerate(SonyVendorId)
        .Select(device => (
            Device: device,
            Product: TryReadProduct(device),
            Manufacturer: TryReadManufacturer(device),
            InputLength: (int)device.Capabilities.InputReportByteLength,
            OutputLength: (int)device.Capabilities.OutputReportByteLength,
            FeatureLength: (int)device.Capabilities.FeatureReportByteLength))
        .Where(device => device.InputLength > 0 && IsPrimaryControllerInterface(device.Device))
        .OrderByDescending(device => ScoreDualSenseLikelihood(device.Product, device.Manufacturer))
        .ThenBy(device => IsBluetoothConnection(device.Device) ? 1 : 0)
        .ThenByDescending(device => device.InputLength)
        .ToList();

    deviceCombo.Items.Clear();
    foreach (var device in devices)
    {
        deviceCombo.Items.Add(FormatDeviceListItem(device));
    }

    if (deviceCombo.Items.Count > 0)
    {
        deviceCombo.SelectedIndex = 0;
        SetStatus("Device ready");
        AppendLog($"Detected {devices.Count} Sony gamepad HID device(s).");
    }
    else
    {
        deviceDetailsLabel.Text = "No Sony USB or Bluetooth gamepad HID devices found.";
        SetStatus("No controller found");
        AppendLog("No Sony USB or Bluetooth gamepad HID devices found. Connect the controller and click Refresh.");
    }

    UpdateActionAvailability();
}

void UpdateSelectedDeviceDetails()
{
    var selected = GetSelectedDevice();
    if (selected is null)
    {
        deviceDetailsLabel.Text = "No device selected.";
        UpdateActionAvailability();
        return;
    }

    deviceDetailsLabel.Text =
        $"PID 0x{selected.Value.Device.Attributes.ProductId:X4} | Product=\"{NullToUnknown(selected.Value.Product)}\" | " +
        $"Input={selected.Value.InputLength} Output={selected.Value.OutputLength} Feature={selected.Value.FeatureLength} | " +
        $"Connection={GetConnectionLabel(selected.Value.Device)} | UsagePage=0x{selected.Value.Device.Capabilities.UsagePage:X2} Usage=0x{selected.Value.Device.Capabilities.Usage:X2}";

    UpdateActionAvailability();
}

void HandleLightbarSliderChanged()
{
    lightbarHue = lightbarSlider.Value;
    var color = ColorFromHue(lightbarHue);
    lightbarColorArgb = color.ToArgb();
    UpdateLightbarPreview();

    if (currentMode == "mapping")
    {
        return;
    }

    var selected = GetSelectedDevice();
    if (selected is null)
    {
        return;
    }

    TryApplyLightbarColor(selected.Value.Device, color, initializeLightbar: true, logErrors: false);
}

void UpdateLightbarPreview()
{
    var color = Color.FromArgb(lightbarColorArgb);
    lightbarPreview.BackColor = color;
    lightbarValueLabel.Text = $"Hue {lightbarHue:D3} deg  RGB {color.R}, {color.G}, {color.B}";
}

(HidDevice Device, string Product, string Manufacturer, int InputLength, int OutputLength, int FeatureLength)? GetSelectedDevice()
{
    var index = deviceCombo.SelectedIndex;
    if (index < 0 || index >= devices.Count)
    {
        return null;
    }

    return devices[index];
}

void UpdateActionAvailability()
{
    if (form.IsDisposed)
    {
        return;
    }

    if (mapButton.InvokeRequired)
    {
        mapButton.BeginInvoke(new Action(UpdateActionAvailability));
        return;
    }

    var hasDevice = GetSelectedDevice() is not null;
    var isBusy = currentMode != "idle";

    refreshButton.Enabled = !isBusy;
    deviceCombo.Enabled = !isBusy;
    mapButton.Enabled = !isBusy && hasDevice;
    stopButton.Enabled = isBusy;
}

void StartRun(string modeName, string statusText, Action<CancellationToken> runAction)
{
    if (activeRunTask is not null && !activeRunTask.IsCompleted)
    {
        return;
    }

    currentMode = modeName;
    activeRunCts = new CancellationTokenSource();
    UpdateActionAvailability();
    SetStatus(statusText);

    activeRunTask = Task.Run(() =>
    {
        try
        {
            runAction(activeRunCts.Token);
        }
        catch (Exception ex) when (IsSpuriousSuccessException(ex))
        {
            AppendLog("Ignored a spurious Win32 success exception from the device API.");
        }
        catch (OperationCanceledException)
        {
            AppendLog("Run cancelled.");
        }
        catch (VigemBusNotFoundException)
        {
            AppendLog("ViGEmBus is not installed, so the virtual Xbox 360 controller cannot be created.");
            AppendLog("This app emulates an Xbox 360 controller, not an Xbox One controller.");
            AppendLog("Install the ViGEmBus driver, then start mapping again.");
            SetStatus("Error");
        }
        catch (Exception ex)
        {
            AppendLog($"Error: {ex.Message}");
            SetStatus("Error");
        }
        finally
        {
            if (!form.IsDisposed)
            {
                form.BeginInvoke(new Action(() =>
                {
                    currentMode = "idle";
                    UpdateActionAvailability();
                    SetStatus("Idle");
                }));
            }
        }
    }, activeRunCts.Token);
}

async Task StopCurrentRunAsync()
{
    if (activeRunCts is null)
    {
        currentMode = "idle";
        UpdateActionAvailability();
        return;
    }

    activeRunCts.Cancel();

    if (activeRunTask is not null)
    {
        try
        {
            await activeRunTask;
        }
        catch
        {
        }
    }

    activeRunTask = null;
    activeRunCts.Dispose();
    activeRunCts = null;
    currentMode = "idle";
    UpdateActionAvailability();
    SetStatus("Idle");
}

void RunMappedMode(HidDevice device, CancellationToken cancellationToken)
{
    OpenDevice(device);
    PrepareDeviceForLowLatency(device);
    var lastLightbarColorArgb = lightbarColorArgb;
    TryApplyLightbarColor(device, Color.FromArgb(lastLightbarColorArgb), initializeLightbar: true, logErrors: true);
    var highResolutionTimerEnabled = TryEnableHighResolutionTimer();
    var previousThreadPriority = Thread.CurrentThread.Priority;
    Thread.CurrentThread.Priority = ThreadPriority.Highest;
    using var client = new ViGEmClient();
    var xbox = client.CreateXbox360Controller();
    xbox.AutoSubmitReport = false;
    ConnectVirtualController(xbox);
    AppendLog("Virtual Xbox 360 controller connected.");

    try
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            if (!device.IsConnected)
            {
                AppendLog("Controller disconnected.");
                break;
            }

            try
            {
                var report = device.ReadReport(ReadTimeoutMs);
                if (report.ReadStatus == HidDeviceData.ReadStatus.WaitTimedOut)
                {
                    continue;
                }

                if (report.ReadStatus == HidDeviceData.ReadStatus.NotConnected)
                {
                    AppendLog("Read stopped: controller is no longer connected.");
                    break;
                }

                if (report.ReadStatus != HidDeviceData.ReadStatus.Success || !report.Exists)
                {
                    AppendLog($"Read status: {report.ReadStatus}");
                    continue;
                }

                var desiredLightbarColorArgb = lightbarColorArgb;
                if (desiredLightbarColorArgb != lastLightbarColorArgb &&
                    TryApplyLightbarColor(device, Color.FromArgb(desiredLightbarColorArgb), initializeLightbar: false, logErrors: false))
                {
                    lastLightbarColorArgb = desiredLightbarColorArgb;
                }

                var raw = report.GetBytes();
                var reportOffset = GetReportOffset(device, raw);
                if (reportOffset < 0 || !RawIndexesExist(raw, reportOffset))
                {
                    continue;
                }

                var leftX = MapUnsignedByteToXboxAxis(raw[LeftStickXByte + reportOffset], invert: false);
                var leftY = MapUnsignedByteToXboxAxis(raw[LeftStickYByte + reportOffset], invert: InvertLeftStickY);
                var rightX = MapUnsignedByteToXboxAxis(raw[RightStickXByte + reportOffset], invert: false);
                var rightY = MapUnsignedByteToXboxAxis(raw[RightStickYByte + reportOffset], invert: InvertRightStickY);
                var leftTrigger = raw[LeftTriggerByte + reportOffset];
                var rightTrigger = raw[RightTriggerByte + reportOffset];

                var buttons1 = raw[Buttons1Byte + reportOffset];
                var buttons2 = raw[Buttons2Byte + reportOffset];
                var buttons3 = raw[Buttons3Byte + reportOffset];
                var dpad = buttons1 & DPadMask;

                ushort xboxButtons = 0;
                unchecked
                {
                    if ((buttons1 & CrossMask) != 0) xboxButtons |= Xbox360Button.A.Value;
                    if ((buttons1 & CircleMask) != 0) xboxButtons |= Xbox360Button.B.Value;
                    if ((buttons1 & SquareMask) != 0) xboxButtons |= Xbox360Button.X.Value;
                    if ((buttons1 & TriangleMask) != 0) xboxButtons |= Xbox360Button.Y.Value;
                    if ((buttons2 & LeftShoulderMask) != 0) xboxButtons |= Xbox360Button.LeftShoulder.Value;
                    if ((buttons2 & RightShoulderMask) != 0) xboxButtons |= Xbox360Button.RightShoulder.Value;
                    if ((buttons2 & LeftThumbMask) != 0) xboxButtons |= Xbox360Button.LeftThumb.Value;
                    if ((buttons2 & RightThumbMask) != 0) xboxButtons |= Xbox360Button.RightThumb.Value;
                    if ((buttons2 & OptionsMask) != 0) xboxButtons |= Xbox360Button.Start.Value;
                    if ((buttons2 & ShareMask) != 0) xboxButtons |= Xbox360Button.Back.Value;
                    // DualSense has more center buttons than an Xbox 360 pad, so PS and touchpad click share Guide.
                    if ((buttons3 & (PsMask | TouchpadMask)) != 0) xboxButtons |= Xbox360Button.Guide.Value;
                    if (dpad == 0 || dpad == 1 || dpad == 7) xboxButtons |= Xbox360Button.Up.Value;
                    if (dpad == 1 || dpad == 2 || dpad == 3) xboxButtons |= Xbox360Button.Right.Value;
                    if (dpad == 3 || dpad == 4 || dpad == 5) xboxButtons |= Xbox360Button.Down.Value;
                    if (dpad == 5 || dpad == 6 || dpad == 7) xboxButtons |= Xbox360Button.Left.Value;
                }

                xbox.SetButtonsFull(xboxButtons);
                xbox.LeftThumbX = leftX;
                xbox.LeftThumbY = leftY;
                xbox.RightThumbX = rightX;
                xbox.RightThumbY = rightY;
                xbox.LeftTrigger = leftTrigger;
                xbox.RightTrigger = rightTrigger;
                xbox.SubmitReport();
            }
            catch (Exception ex) when (IsSpuriousSuccessException(ex))
            {
                continue;
            }
        }
    }
    finally
    {
        Thread.CurrentThread.Priority = previousThreadPriority;
        DisableHighResolutionTimer(highResolutionTimerEnabled);

        try
        {
            xbox.Disconnect();
        }
        catch
        {
        }

        try
        {
            device.CloseDevice();
        }
        catch
        {
        }

        AppendLog("Virtual Xbox 360 controller disconnected.");
    }
}

void OpenDevice(HidDevice device)
{
    if (device.IsOpen)
    {
        return;
    }

    try
    {
        device.OpenDevice(
            DeviceMode.Overlapped,
            DeviceMode.Overlapped,
            (ShareMode)((int)ShareMode.ShareRead | (int)ShareMode.ShareWrite));
    }
    catch (Exception ex) when (IsSpuriousSuccessException(ex) && device.IsOpen)
    {
        return;
    }

    if (!device.IsOpen)
    {
        throw new InvalidOperationException("Failed to open the HID device.");
    }
}

void PrepareDeviceForLowLatency(HidDevice device)
{
    if (!device.IsOpen)
    {
        return;
    }

    var readHandle = device.ReadHandle;
    if (readHandle == IntPtr.Zero || readHandle == new IntPtr(-1))
    {
        return;
    }

    NativeHid.HidD_SetNumInputBuffers(readHandle, PreferredHidInputBufferCount);
    NativeHid.HidD_FlushQueue(readHandle);
}

bool TryApplyLightbarColor(HidDevice device, Color color, bool initializeLightbar, bool logErrors)
{
    var openedHere = false;

    try
    {
        if (!device.IsOpen)
        {
            OpenDevice(device);
            openedHere = true;
        }

        if (initializeLightbar)
        {
            var setupReport = CreateDualSenseLightbarSetupReport(device);
            if (!device.Write(setupReport, ReadTimeoutMs))
            {
                throw new InvalidOperationException("Failed to initialize the DualSense lightbar.");
            }
        }

        var colorReport = CreateDualSenseLightbarColorReport(device, color);
        if (!device.Write(colorReport, ReadTimeoutMs))
        {
            throw new InvalidOperationException("Failed to send the DualSense lightbar color.");
        }

        return true;
    }
    catch (Exception ex) when (IsSpuriousSuccessException(ex))
    {
        return true;
    }
    catch (Exception ex)
    {
        if (logErrors)
        {
            AppendLog($"Lightbar update failed: {ex.Message}");
        }

        return false;
    }
    finally
    {
        if (openedHere)
        {
            try
            {
                device.CloseDevice();
            }
            catch
            {
            }
        }
    }
}

byte[] CreateDualSenseLightbarSetupReport(HidDevice device)
{
    var report = CreateDualSenseOutputReport(device);
    report[GetDualSenseReportByte(device, DualSenseUsbValidFlag2Byte, DualSenseBluetoothValidFlag2Byte)] |= DualSenseLightbarSetupControlEnableFlag;
    report[GetDualSenseReportByte(device, DualSenseUsbLightbarSetupByte, DualSenseBluetoothLightbarSetupByte)] = DualSenseLightbarSetupLightOut;
    FinalizeDualSenseOutputReport(device, report);
    return report;
}

byte[] CreateDualSenseLightbarColorReport(HidDevice device, Color color)
{
    var report = CreateDualSenseOutputReport(device);
    report[GetDualSenseReportByte(device, DualSenseUsbValidFlag1Byte, DualSenseBluetoothValidFlag1Byte)] |= DualSenseLightbarControlEnableFlag;
    report[GetDualSenseReportByte(device, DualSenseUsbLightbarRedByte, DualSenseBluetoothLightbarRedByte)] = color.R;
    report[GetDualSenseReportByte(device, DualSenseUsbLightbarGreenByte, DualSenseBluetoothLightbarGreenByte)] = color.G;
    report[GetDualSenseReportByte(device, DualSenseUsbLightbarBlueByte, DualSenseBluetoothLightbarBlueByte)] = color.B;
    FinalizeDualSenseOutputReport(device, report);
    return report;
}

byte[] CreateDualSenseOutputReport(HidDevice device)
{
    if (IsBluetoothConnection(device))
    {
        var report = new byte[DualSenseOutputReportBluetoothLength];
        report[0] = DualSenseOutputReportBluetoothId;
        report[1] = (byte)((dualSenseOutputSequence & 0x0F) << 4);
        report[2] = DualSenseOutputTag;
        dualSenseOutputSequence = (byte)((dualSenseOutputSequence + 1) & 0x0F);
        return report;
    }

    var usbReport = new byte[DualSenseOutputReportUsbLength];
    usbReport[0] = DualSenseOutputReportUsbId;
    return usbReport;
}

int GetDualSenseReportByte(HidDevice device, int usbIndex, int bluetoothIndex)
{
    return IsBluetoothConnection(device) ? bluetoothIndex : usbIndex;
}

void FinalizeDualSenseOutputReport(HidDevice device, byte[] report)
{
    if (!IsBluetoothConnection(device))
    {
        return;
    }

    var crc = ComputeDualSenseBluetoothCrc(report, report.Length - 4);
    report[^4] = (byte)(crc & 0xFF);
    report[^3] = (byte)((crc >> 8) & 0xFF);
    report[^2] = (byte)((crc >> 16) & 0xFF);
    report[^1] = (byte)((crc >> 24) & 0xFF);
}

uint ComputeDualSenseBluetoothCrc(byte[] report, int lengthWithoutCrc)
{
    var crc = UpdateCrc32(0xFFFFFFFFu, DualSenseOutputCrcSeed);

    for (var i = 0; i < lengthWithoutCrc; i++)
    {
        crc = UpdateCrc32(crc, report[i]);
    }

    return ~crc;
}

uint UpdateCrc32(uint crc, byte value)
{
    crc ^= value;

    for (var i = 0; i < 8; i++)
    {
        crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320u : crc >> 1;
    }

    return crc;
}

void ConnectVirtualController(IXbox360Controller xbox)
{
    try
    {
        xbox.Connect();
    }
    catch (Exception ex) when (IsSpuriousSuccessException(ex))
    {
    }
}

bool RawIndexesExist(byte[] raw, int reportOffset)
{
    return LeftStickXByte + reportOffset < raw.Length &&
           LeftStickYByte + reportOffset < raw.Length &&
           RightStickXByte + reportOffset < raw.Length &&
           RightStickYByte + reportOffset < raw.Length &&
           LeftTriggerByte + reportOffset < raw.Length &&
           RightTriggerByte + reportOffset < raw.Length &&
           Buttons1Byte + reportOffset < raw.Length &&
           Buttons2Byte + reportOffset < raw.Length &&
           Buttons3Byte + reportOffset < raw.Length;
}

bool IsSpuriousSuccessException(Exception ex)
{
    if (ex is Win32Exception win32 && win32.NativeErrorCode == 0)
    {
        return true;
    }

    if (ex is ExternalException && (ex.HResult & 0xFFFF) == 0)
    {
        return true;
    }

    return ex.Message.Contains("operation completed successfully", StringComparison.OrdinalIgnoreCase);
}

bool TryEnableHighResolutionTimer()
{
    try
    {
        return Winmm.timeBeginPeriod(RequestedTimerResolutionMs) == 0;
    }
    catch
    {
        return false;
    }
}

void DisableHighResolutionTimer(bool enabled)
{
    if (!enabled)
    {
        return;
    }

    try
    {
        Winmm.timeEndPeriod(RequestedTimerResolutionMs);
    }
    catch
    {
    }
}

void AppendLog(string message)
{
    if (form.IsDisposed)
    {
        return;
    }

    if (logBox.InvokeRequired)
    {
        logBox.BeginInvoke(new Action<string>(AppendLog), message);
        return;
    }

    logBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
    logBox.SelectionStart = logBox.TextLength;
    logBox.ScrollToCaret();
}

void SetStatus(string text)
{
    if (form.IsDisposed)
    {
        return;
    }

    if (statusLabel.InvokeRequired)
    {
        statusLabel.BeginInvoke(new Action<string>(SetStatus), text);
        return;
    }

    statusLabel.Text = $"Status: {text}";

    var normalized = text.Trim().ToLowerInvariant();
    var (backColor, foreColor) = normalized switch
    {
        "idle" => (Color.FromArgb(235, 240, 247), Color.FromArgb(43, 63, 96)),
        "mapping running" => (Color.FromArgb(228, 244, 234), Color.FromArgb(36, 108, 67)),
        "device ready" => (Color.FromArgb(235, 244, 239), Color.FromArgb(45, 106, 79)),
        "no controller found" => (Color.FromArgb(251, 240, 220), Color.FromArgb(121, 88, 26)),
        "error" => (Color.FromArgb(251, 232, 232), Color.FromArgb(145, 45, 45)),
        _ => (Color.FromArgb(235, 240, 247), Color.FromArgb(43, 63, 96))
    };

    statusLabel.BackColor = backColor;
    statusLabel.ForeColor = foreColor;
}

string FormatDeviceListItem((HidDevice Device, string Product, string Manufacturer, int InputLength, int OutputLength, int FeatureLength) device)
{
    return $"[{GetConnectionLabel(device.Device)}] VID 0x{device.Device.Attributes.VendorId:X4} PID 0x{device.Device.Attributes.ProductId:X4} | {NullToUnknown(device.Product)}";
}

Color ColorFromHue(int hue)
{
    var normalizedHue = ((hue % 360) + 360) % 360;
    var sector = normalizedHue / 60.0;
    var chroma = 1.0;
    var x = chroma * (1.0 - Math.Abs(sector % 2.0 - 1.0));

    var (red, green, blue) = sector switch
    {
        >= 0 and < 1 => (chroma, x, 0.0),
        >= 1 and < 2 => (x, chroma, 0.0),
        >= 2 and < 3 => (0.0, chroma, x),
        >= 3 and < 4 => (0.0, x, chroma),
        >= 4 and < 5 => (x, 0.0, chroma),
        _ => (chroma, 0.0, x)
    };

    return Color.FromArgb(
        (int)Math.Round(red * 255.0),
        (int)Math.Round(green * 255.0),
        (int)Math.Round(blue * 255.0));
}

short MapUnsignedByteToXboxAxis(byte value, bool invert)
{
    var centered = value - 128;
    var scaled = centered * 256;

    if (value == 255)
    {
        scaled = short.MaxValue;
    }
    else if (value == 0)
    {
        scaled = short.MinValue;
    }

    if (invert)
    {
        scaled = -scaled;
    }

    return (short)Math.Clamp(scaled, short.MinValue, short.MaxValue);
}

bool IsPrimaryControllerInterface(HidDevice device)
{
    return device.Capabilities.UsagePage == GenericDesktopUsagePage &&
           (device.Capabilities.Usage == GamepadUsage || device.Capabilities.Usage == JoystickUsage);
}

bool IsBluetoothConnection(HidDevice device)
{
    return device.DevicePath.Contains("BTHENUM", StringComparison.OrdinalIgnoreCase) ||
           device.Capabilities.InputReportByteLength > 64;
}

string GetConnectionLabel(HidDevice device)
{
    return IsBluetoothConnection(device) ? "Bluetooth" : "USB";
}

int GetReportOffset(HidDevice device, byte[] raw)
{
    if (raw.Length == 0)
    {
        return -1;
    }

    if (raw[0] == 0x31)
    {
        return 1;
    }

    if (IsBluetoothConnection(device))
    {
        return -1;
    }

    return 0;
}

int ScoreDualSenseLikelihood(string? product, string? manufacturer)
{
    var score = 0;

    if (!string.IsNullOrWhiteSpace(product) && product.Contains("DualSense", StringComparison.OrdinalIgnoreCase))
    {
        score += 100;
    }

    if (!string.IsNullOrWhiteSpace(product) && product.Contains("Wireless Controller", StringComparison.OrdinalIgnoreCase))
    {
        score += 25;
    }

    if (!string.IsNullOrWhiteSpace(manufacturer) && manufacturer.Contains("Sony", StringComparison.OrdinalIgnoreCase))
    {
        score += 10;
    }

    return score;
}

string TryReadProduct(HidDevice device)
{
    try
    {
        return device.ReadProduct(out var data) ? DecodeUsbString(data) : string.Empty;
    }
    catch
    {
        return string.Empty;
    }
}

string TryReadManufacturer(HidDevice device)
{
    try
    {
        return device.ReadManufacturer(out var data) ? DecodeUsbString(data) : string.Empty;
    }
    catch
    {
        return string.Empty;
    }
}

string DecodeUsbString(byte[] data)
{
    return Encoding.Unicode.GetString(data).TrimEnd('\0', ' ');
}

string NullToUnknown(string? value)
{
    return string.IsNullOrWhiteSpace(value) ? "unknown" : value;
}

static class Winmm
{
    [DllImport("winmm.dll")]
    internal static extern uint timeBeginPeriod(uint uPeriod);

    [DllImport("winmm.dll")]
    internal static extern uint timeEndPeriod(uint uPeriod);
}

static class NativeHid
{
    [DllImport("hid.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool HidD_FlushQueue(IntPtr hidDeviceObject);

    [DllImport("hid.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool HidD_SetNumInputBuffers(IntPtr hidDeviceObject, int numberBuffers);
}
