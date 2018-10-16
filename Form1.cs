using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Management;
using System.Timers;
using System.Windows.Forms;
using Timer = System.Timers.Timer;

namespace Hyper_V_Manager
{
    /// <summary>
    /// Possible VM States
    /// </summary>
    public enum VmState
    {
        Unknown = 0,
        Other = 1,          // Corresponds to CIM_EnabledLogicalElement.EnabledState = Other.
        Running = 2,        // Enabled
        Stopped = 3,        // Disabled
        ShutDown = 4,       // Valid in version 1 (V1) of Hyper-V only. The virtual machine is shutting down via the shutdown service. Corresponds to CIM_EnabledLogicalElement.EnabledState = ShuttingDown.
        Saved = 6,          // Corresponds to CIM_EnabledLogicalElement.EnabledState = Enabled but offline.
        Paused = 9,         // Corresponds to CIM_EnabledLogicalElement.EnabledState = Quiesce, Enabled but paused.
        Starting = 32770,
        Saving = 32773,
        Stopping = 32774,
        Pausing = 32776,
        Resuming = 32777
    }

    /// <inheritdoc />
    /// <summary>
    /// Main Form
    /// </summary>
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private readonly Timer _timer = new Timer();
        private readonly Dictionary<string, string> _changingVMs = new Dictionary<string, string>();

        /// <summary>
        /// Form load event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            _timer.Elapsed += TimerElapsed;
            _timer.Interval = 4500;
            BuildContextMenu();
        }

        /// <summary>
        /// Time elapsed event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            if(!contextMenuStrip1.Visible)
                UpdateBalloontip();
        }

        /// <summary>
        /// Update Balloon Tooltip on VM State Change
        /// </summary>
        private void UpdateBalloontip()
        {
            var localVMs = GetVMs();

            foreach (var vm in localVMs)
            {
                if (!_changingVMs.ContainsKey(vm["ElementName"].ToString())) continue;
                var initvmBalloonState = _changingVMs[vm["ElementName"].ToString()];
                var vmState = (VmState) Convert.ToInt32(vm["EnabledState"]);
                var currentBalloonState = vmState.ToString();

                if (initvmBalloonState != currentBalloonState)
                {
                    notifyIcon1.ShowBalloonTip(4000, "VM State Changed", vm["ElementName"] + " " + currentBalloonState, ToolTipIcon.Info);
                    _changingVMs[vm["ElementName"].ToString()] = currentBalloonState;
                }
                else if (vmState == VmState.Running || vmState == VmState.Stopped || vmState == VmState.Paused || vmState == VmState.Saved)
                    _changingVMs.Remove(vm["ElementName"].ToString());
                else if (_changingVMs.Count <= 0)
                    _timer.Enabled = false;
            }
        }

        /// <summary>
        /// Get an enumerable list of Hyper-V VMs using WMI
        /// </summary>
        /// <returns></returns>
        private static IEnumerable<ManagementObject> GetVMs()
        {
            var vms = new List<ManagementObject>();

            var path = ConfigurationManager.AppSettings["root"];
            var manScope = new ManagementScope(path);
            var queryObj = new ObjectQuery("Select * From Msvm_ComputerSystem");

            var vmSearcher = new ManagementObjectSearcher(manScope, queryObj);
            var vmCollection = vmSearcher.Get();

            foreach (var o in vmCollection)
            {
                var vm = (ManagementObject) o;
                if (vm.Properties["Caption"].Value.ToString() == "Virtual Machine")
                {
                    vms.Add(vm);
                }
            }

            return vms;
        }

        #region ContextMenuEvents

        private void VmItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start($@"{Environment.GetEnvironmentVariable("SYSTEMROOT")}\System32\vmconnect.exe",
                $"localhost \"{((ToolStripMenuItem) sender).Name}\"");
        }


        private void PauseItem_Click(object sender, EventArgs e)
        {
            ChangeVmState("Pause", ((ToolStripMenuItem)sender).OwnerItem.Name);
        }

        private void SaveStateItem_Click(object sender, EventArgs e)
        {
            ChangeVmState("Save State", ((ToolStripMenuItem)sender).OwnerItem.Name);
        }

        private void ShutDownItem_Click(object sender, EventArgs e)
        {
            ChangeVmState("Shut Down", ((ToolStripMenuItem)sender).OwnerItem.Name);
        }

        private void StopItem_Click(object sender, EventArgs e)
        {
            ChangeVmState("Stop", ((ToolStripMenuItem)sender).OwnerItem.Name);
        }

        private void StartItem_Click(object sender, EventArgs e)
        {
            ChangeVmState("Start", ((ToolStripMenuItem)sender).OwnerItem.Name);
        }

        private void ExitItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        #endregion


        /// <summary>
        /// Set the VM State
        /// </summary>
        /// <param name="requestedState">Requested state</param>
        /// <param name="vmName">Name of the VM</param>
        private void ChangeVmState(string requestedState, string vmName)
        {
            var localVMs = GetVMs();
            
            _timer.Enabled = true;

            // Loop to find a matching VM
            foreach (var vm in localVMs)
            {
                if (vm["ElementName"].ToString() != vmName) continue;

                // Set the state to unknown as we request the change
                _changingVMs[vm["ElementName"].ToString()] = VmState.Unknown.ToString();

                var inParams = vm.GetMethodParameters("RequestStateChange");

                switch (requestedState)
                {
                    case "Start":
                        inParams["RequestedState"] = (ushort) VmState.Running;
                        break;
                    case "Stop":
                        inParams["RequestedState"] = (ushort) VmState.Stopped;
                        break;
                    case "Shut Down":
                        inParams["RequestedState"] = (ushort) VmState.ShutDown;
                        break;
                    case "Pause":
                        inParams["RequestedState"] = (ushort) VmState.Paused;
                        break;
                    case "Save State":
                        inParams["RequestedState"] = (ushort) VmState.Saved;
                        break;
                    default:
                        throw new Exception("Unexpected VM State");
                }
                // Todo - handle response from request to change
                // https://docs.microsoft.com/en-us/windows/desktop/hyperv_v2/requeststatechange-msvm-computersystem
                vm.InvokeMethod("RequestStateChange", inParams, null);
            }
        }

        /// <summary>
        /// Build context menu
        /// </summary>
        private void BuildContextMenu()
        {
            // Get all VMs
            var localVMs = GetVMs();

            // Clear the context menu
            contextMenuStrip1.Items.Clear();

            // Add context menu items with current state options
            foreach (var vm in localVMs)
            {
                var startItem = new ToolStripMenuItem("Start");
                startItem.Click += StartItem_Click;
                startItem.DisplayStyle = ToolStripItemDisplayStyle.Text;

                var stopItem = new ToolStripMenuItem("Stop");
                stopItem.Click += StopItem_Click;
                stopItem.DisplayStyle = ToolStripItemDisplayStyle.Text;

                var shutDownItem = new ToolStripMenuItem("Shut Down");
                shutDownItem.Click += ShutDownItem_Click;
                shutDownItem.DisplayStyle = ToolStripItemDisplayStyle.Text;

                var saveStateItem = new ToolStripMenuItem("Save State");
                saveStateItem.Click += SaveStateItem_Click;
                saveStateItem.DisplayStyle = ToolStripItemDisplayStyle.Text;

                var pauseItem = new ToolStripMenuItem("Pause");
                pauseItem.Click += PauseItem_Click;
                pauseItem.DisplayStyle = ToolStripItemDisplayStyle.Text;

                var vmState = (VmState)Convert.ToInt32(vm["EnabledState"]);

                // ReSharper disable once SwitchStatementMissingSomeCases
                switch (vmState)
                {
                    case VmState.Running:
                        startItem.Enabled = false;
                        break;
                    case VmState.Saved:
                    case VmState.ShutDown:
                    case VmState.Stopped:
                        stopItem.Enabled = false;
                        saveStateItem.Enabled = false;
                        pauseItem.Enabled = false;
                        break;
                    case VmState.Paused:
                        pauseItem.Enabled = false;
                        break;
                }

                // Create a lable for each VM, show its state if not stopped
                var vmStatusText = vm["ElementName"].ToString();
                if (vmState != VmState.Stopped) vmStatusText += " [" + vmState + "]";

                // Create sub-menu
                var vmItem = new ToolStripMenuItem(vmStatusText) {Name = vm["ElementName"].ToString()};

                // Add a VM click handler to open remote
                vmItem.Click += VmItem_Click;

                // Add sub-menu items
                if (vmState == VmState.Running || vmState == VmState.Stopped || vmState == VmState.Saved || vmState == VmState.Paused)
                {
                    vmItem.DropDownItems.Add(startItem);
                    vmItem.DropDownItems.Add(stopItem);
                    vmItem.DropDownItems.Add(shutDownItem);
                    vmItem.DropDownItems.Add(saveStateItem);
                    vmItem.DropDownItems.Add(pauseItem);
                }
                contextMenuStrip1.Items.Add(vmItem);
            }

            // Add Exit option
            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += ExitItem_Click;
            contextMenuStrip1.Items.Add(exitItem);

            // Redraw the menu
            contextMenuStrip1.Refresh();
        }

        /// <summary>
        /// Context menu open event
        /// Build a new context menu op open
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ContextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            BuildContextMenu();       
        }

        /// <summary>
        /// Form activated event
        /// Hide the form, to only show the system tray
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Activated(object sender, EventArgs e)
        {
            Hide();
        }

    }
}
