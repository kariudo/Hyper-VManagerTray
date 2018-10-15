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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /*
         * 
         * Unknown      - 0     - state could not be determined
         * Enabled      - 2     - VM is running
         * Disabled     - 3     - VM is stopped
         * Paused       - 32768 - VM is paused
         * Suspended    - 32769 - VM is in a saved state
         * Starting     - 32770 - VM is starting
         * Saving       - 32773 - VM is saving its state
         * Stopping     - 32774 - VM is turning off
         * Pausing      - 32776 - VM is pausing
         * Resuming     - 32777 - VM is resuming from a paused state
         * 
         */


        private readonly Timer _timer = new Timer();
        private readonly Dictionary<string, string> _changingVMs = new Dictionary<string, string>();

        
        private void Form1_Load(object sender, EventArgs e)
        {
            _timer.Elapsed += TimerElapsed;
            _timer.Interval = 4500;
            BuildContextMenu();
        }

        private void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            if(!contextMenuStrip1.Visible)
                UpdateBalloontip();
        }

        
        private void UpdateBalloontip()
        {
            var localVMs = GetVMs();

            foreach (var vm in localVMs)
            {
                if (!_changingVMs.ContainsKey(vm["ElementName"].ToString())) continue;
                string currentBalloonState;
                var initvmBalloonState = _changingVMs[vm["ElementName"].ToString()];

                switch (Convert.ToInt32(vm["EnabledState"]))
                {
                    case 2:
                        currentBalloonState = "Running";
                        break;
                    case 3:
                        currentBalloonState = "Stopped";
                        break;
                    case 32768:
                        currentBalloonState = "Paused";
                        break;
                    case 32769:
                        currentBalloonState = "Saved";
                        break;
                    case 32770:
                        currentBalloonState = "Starting";
                        break;
                    case 32773:
                        currentBalloonState = "Saving";
                        break;
                    case 32774:
                        currentBalloonState = "Stopping";
                        break;
                    default:
                        currentBalloonState = "Unknown";
                        break;
                }

                if (initvmBalloonState != currentBalloonState)
                {
                    notifyIcon1.ShowBalloonTip(4000, "VM State Changed", vm["ElementName"] + " " + currentBalloonState, ToolTipIcon.Info);
                    _changingVMs[vm["ElementName"].ToString()] = currentBalloonState;
                }
                else if (currentBalloonState == "Running" || currentBalloonState == "Stopped" || currentBalloonState == "Paused" || currentBalloonState == "Saved")
                    _changingVMs.Remove(vm["ElementName"].ToString());
                else if (_changingVMs.Count <= 0)
                    _timer.Enabled = false;
            }
        }



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

        /*
         * Context menu events
         */
        #region ContextMenuEvents


        void PauseItem_Click(object sender, EventArgs e)
        {
            ChangeVmState("Pause", ((ToolStripMenuItem)sender).OwnerItem.Name);
        }

        void SaveStateItem_Click(object sender, EventArgs e)
        {
            ChangeVmState("Save State", ((ToolStripMenuItem)sender).OwnerItem.Name);
        }

        void StopItem_Click(object sender, EventArgs e)
        {
            ChangeVmState("Stop", ((ToolStripMenuItem)sender).OwnerItem.Name);
        }

        void StartItem_Click(object sender, EventArgs e)
        {
            ChangeVmState("Start", ((ToolStripMenuItem)sender).OwnerItem.Name);
        }

        void ExitItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        #endregion


        /*
         * Change the state of the VM based on the state passed in and the VM Name
         */
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
                        inParams["RequestedState"] = 2;
                        break;
                    case "Stop":
                        inParams["RequestedState"] = 3;
                        break;
                    case "Pause":
                        inParams["RequestedState"] = 32768;
                        break;
                    case "Save State":
                        inParams["RequestedState"] = 32769;
                        break;
                    default:
                        throw new Exception("Unexpected VM State");
                }

                vm.InvokeMethod("RequestStateChange", inParams, null);
            }
        }


        /*
         * Build out the context menu items.
         */


        private void BuildContextMenu()
        {
            //Re-get the VM's from Hyper-V
            var localVMs = GetVMs();

            //Clear out the menu items
            contextMenuStrip1.Items.Clear();

            //loop through the VMs and rebuild out the context menus
            //Do this to ensure the proper status is displayed and the proper action items are enabled
            foreach (var vm in localVMs)
            {

                string vmState;

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

                switch (Convert.ToInt32(vm["EnabledState"]))
                {
                    case 2:
                        vmState = "Running";
                        startItem.Enabled = false;
                        break;
                    case 3:
                        vmState = "Stopped";
                        stopItem.Enabled = false;
                        saveStateItem.Enabled = false;
                        pauseItem.Enabled = false;
                        break;
                    case 32768:
                        vmState = "Paused";
                        pauseItem.Enabled = false;
                        break;
                    case 32769:
                        vmState = "Saved";
                        saveStateItem.Enabled = false;
                        stopItem.Enabled = false;
                        pauseItem.Enabled = false;
                        break;
                    case 32770:
                        vmState = "Starting";
                        break;
                    case 32773:
                        vmState = "Saving";
                        break;
                    case 32774:
                        vmState = "Stopping";
                        break;
                    default:
                        vmState = "Unknown";
                        break;
                }

                var vmStatusText = vm["ElementName"] + " " + vmState;

                var vmItem = new ToolStripMenuItem(vmStatusText) {Name = vm["ElementName"].ToString()};


                if (vmState == "Running" || vmState == "Stopped" || vmState == "Saved" || vmState == "Paused")
                {
                    vmItem.DropDownItems.Add(startItem);
                    vmItem.DropDownItems.Add(stopItem);
                    vmItem.DropDownItems.Add(saveStateItem);
                    vmItem.DropDownItems.Add(pauseItem);
                }
            
            
                contextMenuStrip1.Items.Add(vmItem);
                
            }

            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += ExitItem_Click;
            contextMenuStrip1.Items.Add(exitItem);
            contextMenuStrip1.Refresh();
        }

        /*
         * When the context menu is opened, rebuild the Context Menu Items.  This ensures the status and items are up to date
         */
        private void ContextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            BuildContextMenu();       
        }

        /*
         *  When the form is activated, hide it so the only UI is the system tray icon.
         */
        private void Form1_Activated(object sender, EventArgs e)
        {
            Hide();
        }

    }
}
