# Hyper-VManagerTray
Provides access to Hyper-V (V2) Management from the system tray

## Requirements
  * Hyper-V (V2+)
  * .NET 4.7.2

## Usage
Just run the `.exe`.

## Features
 * Right-click to access a list of available VMs on the host
 * The state of any VM not "Stopped" will be shown as well in the list with its name
 * Clicking on an entry in the list will open `vmconnect` to that vm
 * Sub-menu provides access to the following state options:
   * Start
   * Stop
   * Shut Down
   * Save State
   * Pause
 * Notifications will be popped for state changes


## Credits
The original idea and basis for this tool was from [Jerry Ormans Microsoft blog article](https://blogs.msdn.microsoft.com/jorman/2010/01/24/hyper-v-manager/);
however, his tool had stopped working for everyone years ago, and I needed one that worked.
As such this is a rewrite based on that idea, but I would like to thank him for the (though
non-working) starting point. Aesthetically its very similar with the addition of a number of
features and fixes, as well as being updated to use current WMI.
