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
        Running = 2,
        Stopped = 3,
        Paused = 32768,
        Saved = 32769,
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


        private void PauseItem_Click(object sender, EventArgs e)
        {
            ChangeVmState("Pause", ((ToolStripMenuItem)sender).OwnerItem.Name);
        }

        private void SaveStateItem_Click(object sender, EventArgs e)
        {
            ChangeVmState("Save State", ((ToolStripMenuItem)sender).OwnerItem.Name);
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
        /// <param name="vmState">Requested state</param>
        /// <param name="vmName">Name of the VM</param>
        private void ChangeVmState(string vmState, string vmName)
        {
            var localVMs = GetVMs();
            
            _timer.Enabled = true;

            foreach (var vm in localVMs)
            {
                if (vm["ElementName"].ToString() != vmName) continue;
                _changingVMs.Add(vm["ElementName"].ToString(), "Unknown");

                var inParams = vm.GetMethodParameters("RequestStateChange");

                switch (vmState)
                {
                    case "Start":
                        inParams["RequestedState"] = (int) VmState.Running;
                        break;
                    case "Stop":
                        inParams["RequestedState"] = (int) VmState.Stopped;
                        break;
                    case "Pause":
                        inParams["RequestedState"] = (int) VmState.Paused;
                        break;
                    case "Save State":
                        inParams["RequestedState"] = (int) VmState.Saved;
                        break;
                    default:
                        throw new Exception("Unexpected VM State");
                }

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
                    case VmState.Stopped:
                        stopItem.Enabled = false;
                        saveStateItem.Enabled = false;
                        pauseItem.Enabled = false;
                        break;
                    case VmState.Paused:
                        pauseItem.Enabled = false;
                        break;
                    case VmState.Saved:
                        saveStateItem.Enabled = false;
                        stopItem.Enabled = false;
                        pauseItem.Enabled = false;
                        break;
                }

                // Create a lable for each VM, show its state if not stopped
                var vmStatusText = vm["ElementName"].ToString();
                if (vmState != VmState.Stopped) vmStatusText += " [" + vmState + "]";

                // Create sub-menu
                var vmItem = new ToolStripMenuItem(vmStatusText) {Name = vm["ElementName"].ToString()};

                // Add sub-menu items
                if (vmState == VmState.Running || vmState == VmState.Stopped || vmState == VmState.Saved || vmState == VmState.Paused)
                {
                    vmItem.DropDownItems.Add(startItem);
                    vmItem.DropDownItems.Add(stopItem);
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
