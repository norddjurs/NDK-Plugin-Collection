using System;
using NDK.Framework;
using System.Data;
using System.DirectoryServices.AccountManagement;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;

namespace NDK.PluginCollection {

	#region SendWindowsMessage class.
	public class SendWindowsMessage : BasePlugin {

		#region Implement PluginBase abstraction.
		/// <summary>
		/// Gets the unique plugin guid.
		/// When implementing a plugin, this method should return the same unique guid every time. 
		/// </summary>
		/// <returns></returns>
		public override Guid GetGuid() {
			return new Guid("{0A5F7B2F-E4C3-41C3-8FE3-164112F8B6AD}");
		} // GetGuid

		/// <summary>
		/// Gets the the plugin name.
		/// When implementing a plugin, this method should return a proper display name.
		/// </summary>
		/// <returns></returns>
		public override String GetName() {
			return "NDK Send Windows Message";
		} // GetName

		/// <summary>
		/// Run the plugin.
		/// When implementing a plugin, this method is invoked by the service application or the commandline application.
		/// 
		/// If the method finishes when invoked by the service application, it is reinvoked after a short while as long as the
		/// service application is running.
		/// 
		/// Take care to write good comments in the code, log as much as possible, as correctly as possible (little normal, much debug).
		/// </summary>
		public override void Run() {
			try {
				// Get configurations.
				Int32 timeoutSeconds = this.GetLocalValue("Timeout", 30);
				DateTime timeout = DateTime.Now.AddSeconds(timeoutSeconds);
				String processName = this.GetLocalValue("ProcessName", String.Empty);
				String processClass = this.GetLocalValue("ProcessClass", String.Empty);

				Boolean logProcesses = false;
				Boolean logClasses = false;

				// Keep retrying, untill the timeout.
				while (timeout.CompareTo(DateTime.Now) > 0) {
					// List process names.
					if (logProcesses == false) {
						logProcesses = true;
						foreach (Process process in Process.GetProcesses()) {
							this.Log("Process: {0}, {1}", process.Id, process.ProcessName);
						}
					}

					// Get the process, identified by the process name.
					if (processName.Trim().Length > 0) {
						Process[] processes = Process.GetProcessesByName(processName);
						Process process = null;
						Int32 processIndex = 0;
						while ((processIndex < processes.Length) && (process == null)) {
							process = processes[processIndex];
						}

						if (process != null) {
							// List window handles for the process.
							if (logClasses == false) {
								logClasses = true;
								Int32 maxCounter = 1000;
								IntPtr previousChildHandle = IntPtr.Zero;
								IntPtr currentChildHandle = IntPtr.Zero;
								StringBuilder currentChildClassName = new StringBuilder(256);
								while ((true) && (maxCounter > 0)) {
									try {
										// Get the current child handle.
										currentChildHandle = FindWindowEx(process.MainWindowHandle, previousChildHandle, null, null);
										if (currentChildHandle == IntPtr.Zero) {
											break;
										}

										// Iterate.
										previousChildHandle = currentChildHandle;
										maxCounter--;

										// Get the class name of the current child handle.
										GetClassName(currentChildHandle, currentChildClassName, currentChildClassName.Capacity);
										this.Log("Handle: {0}, {1}", currentChildHandle, currentChildClassName);
									} catch { }
								}
							}


							// Send the Windows Message.
							IntPtr child = FindWindowEx(process.MainWindowHandle, new IntPtr(0), processClass, null);
							SendMessage(child, 0x000C, 0, "Hello Notepad World!" + Environment.NewLine);
							this.Log("Sending Windows Message");

							// Exit on success.
							break;
						} else {
							// Wait.
							Thread.Sleep(200);
						}
					}
				}
			} catch (Exception exception) {
				// Send message on error.
				this.SendMail("Error " + this.GetName(), exception.Message, false);

				// Throw the error.
				throw;
			}
		} // Run

		[DllImport("user32.dll")]
		static extern Int32 GetClassName(IntPtr hWnd, StringBuilder lpClassName, Int32 nMaxCount);

		[DllImport("user32.dll")]
		public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, String lpszClass, String lpszWindow);

		[DllImport("User32.dll")]
		public static extern Int32 SendMessage(IntPtr hWnd, Int32 uMsg, Int32 wParam, String lParam);

		[DllImport("User32.dll", EntryPoint = "SendMessage")]
		public static extern Int32 SendMessage1(IntPtr hWnd, Int32 uMsg, Int32 wParam, IntPtr lParam);

		[DllImport("user32.dll")]
		private static extern Boolean PostMessage(IntPtr hwnd, Int32 msg, IntPtr wparam, IntPtr lparam);

		[DllImport("user32.dll")]
		private static extern Int32 RegisterWindowMessage(String message);

		#endregion

	} // SendWindowsMessage
	#endregion

} // NDK.PluginCollection