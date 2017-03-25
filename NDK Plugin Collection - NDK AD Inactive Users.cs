using System;
using NDK.Framework;
using System.Data;
using System.DirectoryServices.AccountManagement;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace NDK.PluginCollection {

	#region ActiveDirectoryInactiveUsers class.
	public class ActiveDirectoryInactiveUsers : BasePlugin {

		#region Implement PluginBase abstraction.
		/// <summary>
		/// Gets the unique plugin guid.
		/// When implementing a plugin, this method should return the same unique guid every time. 
		/// </summary>
		/// <returns></returns>
		public override Guid GetGuid() {
			return new Guid("{9EA6C7E2-4822-4F01-A02F-AEC7C4AC84C3}");
		} // GetGuid

		/// <summary>
		/// Gets the the plugin name.
		/// When implementing a plugin, this method should return a proper display name.
		/// </summary>
		/// <returns></returns>
		public override String GetName() {
			return "NDK ActiveDirectory Inactive Users";
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
				// Get configuration values.
				Boolean configMessageSend = this.GetLocalValue("MessageSend", true);
				List<String> configMessageTo = this.GetLocalValues("MessageTo");
				String configMessageSubject = this.GetLocalValue("MessageSubject", this.GetName());
				Boolean configFailAlways = this.GetLocalValue("FailAlways", true);
				String configBaseDN = this.GetLocalValue("BaseDN", String.Empty).Trim();
				String configInfoText = this.GetLocalValue("InfoText", "User automatically disabled.").Trim();
				Int32 configInactivePeriodDays = this.GetLocalValue("InactivePeriodDays", 90);
				String configInactiveAction = this.GetLocalValue("InactiveAction", "DISABLE").ToLower();

				// Get the inactive date.
				if (configInactivePeriodDays < 1) {
					configInactivePeriodDays = 90;
				}
				DateTime configInactivePeriod = DateTime.Now.AddDays(-configInactivePeriodDays);

				// Collect the inactive users, that validates according to the configured values.
				List<AdUser> inactiveUsers = new List<AdUser>();

				// Collect errors, when performing the desited action.
				List<String> errors = new List<String>();

				// Initialize the user validator.
				ActiveDirectoryUserValidator userValidator = new ActiveDirectoryUserValidator(this);

				// Log.
				this.Log("Processing inactive users for {0} days, since {1:yyyy-MM-dd}.", configInactivePeriodDays, configInactivePeriod);

				// Get all inactive users.
				foreach (AdUser user in this.GetAllUsers(UserQuery.ALL).OrderBy(x => x.LastLogon)) {
					if ((user.LastLogon != null) && (user.LastLogon.Value.CompareTo(configInactivePeriod) < 0)) {
						Boolean userIsValidated = false;
						Boolean userInBaseDN = false;

						// Log.
						this.Log("Inactive user: {0} - {1} ({2:yyyy-MM-dd})", user.SamAccountName, user.Name, user.LastLogon);

						// Validate the user, according to the configured values.
						if (userValidator.ValidateUser(user) == true) {
							// Success.
							userIsValidated = true;
						} else {
							// Fail.
							userIsValidated = false;

							// Log.
							this.Log("The user is not applicable, according to the configured values");
						}

						// Validate that the user is below the base DN.
						if ((configBaseDN.Length == 0) || ((configBaseDN.Length > 0) && (user.DistinguishedName.EndsWith(configBaseDN) == true))) {
							// Success.
							userInBaseDN = true;
						} else {
							// Fail.
							userInBaseDN = false;

							// Log.
							this.Log("The user is not below the base DN '{0}'. User DN is '{1}'.", configBaseDN, user.DistinguishedName);
						}

						// Add the user.
						if ((userIsValidated == true) && (userInBaseDN == true)) {
							inactiveUsers.Add(user);
						}
					}
				}

				// Process found inactive users.
				if (configFailAlways == false) {
					foreach (AdUser user in inactiveUsers) {
						switch (configInactiveAction) {
							case "disable":
									// Log.
									this.Log("Disabling inactive user: {0} - {1} ({2:yyyy-MM-dd})", user.SamAccountName, user.Name, user.LastLogon);

								// Disable the user.
								try {
									//user.Enabled = false;
									//user.Info = String.Format("{0:yyyy-MM-dd} {0}", DateTime.Now, configInfoText) + Environment.NewLine + user.Info;
									//user.Save();
								} catch (Exception exception) {
									errors.Add(String.Format("Unable to disabling inactive user: {0} - {1}: ({2})", user.SamAccountName, user.Name, exception.Message));

									// Log.
									this.LogError(exception);
								}
								break;
							case "delete":
								// Log.
								this.Log("Deleting inactive user: {0} - {1} ({2:yyyy-MM-dd})", user.SamAccountName, user.Name, user.LastLogon);

								// Disable the user.
								// The user is disabled before it is deleted, incase it can not be deleted.
								try {
									//user.Enabled = false;
									//user.Info = String.Format("{0:yyyy-MM-dd} {0}", DateTime.Now, configInfoText) + Environment.NewLine + user.Info;
									//user.Save();
								} catch (Exception exception) {
									errors.Add(String.Format("Unable to disabling inactive user: {0} - {1}: ({2})", user.SamAccountName, user.Name, exception.Message));

									// Log.
									this.LogError(exception);
								}

								// Delete the user.
								try {
									//user.Delete();
								} catch (Exception exception) {
									errors.Add(String.Format("Unable to delete inactive user: {0} - {1}: ({2})", user.SamAccountName, user.Name, exception.Message));

									// Log.
									this.LogError(exception);
								}
								break;
							default:
								break;
						}
					}
				} else {
					// Log.
					this.Log("None of the inactive users was processed, because the configuration 'FailAlways' is enabled.");
				}

				// Send message.
				if (configMessageSend == true) {
					if (configMessageTo.Count > 0) {
						this.SendMail(
							String.Join(";", configMessageTo.ToArray()),
							configMessageSubject,
							this.GetHtmlMessage(configFailAlways, configBaseDN, configInactivePeriodDays, configInactiveAction, userValidator, inactiveUsers, errors),
							true
						);
					} else {
						this.SendMail(
							configMessageSubject,
							this.GetHtmlMessage(configFailAlways, configBaseDN, configInactivePeriodDays, configInactiveAction, userValidator, inactiveUsers, errors),
							true
						);
					}
				}
			} catch (Exception exception) {
				// Send message on error.
				this.SendMail("Error " + this.GetName(), exception.Message, false);

				// Throw the error.
				throw;
			}
		} // Run

		private String GetHtmlMessage(Boolean configFailAlways, String configBaseDN, Int32 configInactivePeriodDays, String configInactiveAction, ActiveDirectoryUserValidator userValidator, List<AdUser> inactiveUsers, List<String> errors) {
			DateTime configInactivePeriod = DateTime.Now.AddDays(-configInactivePeriodDays);
			StringBuilder html = new StringBuilder();

			html.AppendLine(@"<!DOCTYPE html>");
			html.AppendLine(@"<html>");
			html.AppendLine(@"	<head>");
			html.AppendLine(@"		<meta  content=""text /html; charset=UTF-8""  http-equiv=""content-type"">");
			html.AppendLine(@"		<title>Advis</title>");
			html.AppendLine(@"		<style>");
			html.AppendLine(@"			h2 {");
			html.AppendLine(@"				color: navy;");
			html.AppendLine(@"			} ");
			html.AppendLine(@"			th {");
			html.AppendLine(@"				background-color: #72D2FF;");
			html.AppendLine(@"				text-align: left;");
			html.AppendLine(@"				vertical-align: top;");
			html.AppendLine(@"			}");
			html.AppendLine(@"			td {");
			html.AppendLine(@"				text-align: left;");
			html.AppendLine(@"				vertical-align: top;");
			html.AppendLine(@"			}");
			html.AppendLine(@"		</style>");
			html.AppendLine(@"	</head>");
			html.AppendLine(@"	<body>");

			// Message.
			if (configFailAlways == false) {
				html.AppendLine(@"		<p>This automatic task has performed the specified action on the inactive users in the active directory.");
				html.AppendLine(@"		</p>");
			} else {
				html.AppendLine(@"		<p>This automatic task is DEACTIVATED.<br>");
				html.AppendLine(@"		<p>When it is enabled, it will perform the specified action on the inactive users in the active directory.");
				html.AppendLine(@"		</p>");
			}

			// Inactive users.
			html.AppendLine(@"		<h2>Inactive users</h2>");
			html.AppendLine(@"		<table border=""0"" cellpadding=""5"" cellspacing=""0"">");
			html.AppendLine(@"			<thead>");
			html.AppendLine(@"				<tr>");
			html.AppendLine(@"					<th>");
			html.AppendLine($"						Userid<br>");
			html.AppendLine(@"					</th>");

			html.AppendLine(@"					<th>");
			html.AppendLine($"						Full name<br>");
			html.AppendLine(@"					</th>");

			html.AppendLine(@"					<th>");
			html.AppendLine($"						E-mail<br>");
			html.AppendLine(@"					</th>");

			html.AppendLine(@"					<th>");
			html.AppendLine($"						Last logon<br>");
			html.AppendLine(@"					</th>");
			html.AppendLine(@"				</tr>");
			html.AppendLine(@"			</thead>");
			html.AppendLine(@"			<tbody>");

			foreach (AdUser user in inactiveUsers) {
				html.AppendLine(@"				<tr>");
				html.AppendLine(@"					<td>");
				html.AppendLine($"						{user.SamAccountName}<br>");
				html.AppendLine(@"					</td>");

				html.AppendLine(@"					<td>");
				html.AppendLine($"						{user.Name}<br>");
				html.AppendLine(@"					</td>");

				html.AppendLine(@"					<td>");
				html.AppendLine($"						{user.EmailAddress}<br>");
				html.AppendLine(@"					</td>");

				html.AppendLine(@"					<td>");
				html.AppendLine($"						{user.LastLogon.Value:yyyy-MM-dd}<br>");
				html.AppendLine(@"					</td>");
				html.AppendLine(@"				</tr>");
			}
			html.AppendLine(@"			</tbody>");
			html.AppendLine(@"			<tfoot>");
			html.AppendLine(@"				<tr>");
			html.AppendLine(@"					<td colspan=""4"">");
			html.AppendLine($"						{inactiveUsers.Count} inactive users");
			html.AppendLine(@"					</td>");
			html.AppendLine(@"				</tr>");
			html.AppendLine(@"			</tfoot>");
			html.AppendLine(@"		</table>");

			html.AppendLine(@"		<h2>Configuration</h2>");
			html.AppendLine(@"		<table border=""0"" cellpadding=""5"" cellspacing=""0"">");
			html.AppendLine(@"			<tbody>");

			// Period.
			html.AppendLine(@"				<tr>");
			html.AppendLine(@"					<th>Inactive period</th>");
			html.AppendLine(@"					<td>");
			html.AppendLine($"						{configInactivePeriodDays} days, since {configInactivePeriod:yyyy-MM-dd}<br>");
			html.AppendLine(@"					</td>");
			html.AppendLine(@"				</tr>");

			// Action.
			html.AppendLine(@"				<tr>");
			html.AppendLine(@"					<th>Inactive action</th>");
			html.AppendLine(@"					<td>");
			html.AppendLine($"						{configInactiveAction.ToUpper()} the inactive users<br>");
			html.AppendLine(@"					</td>");
			html.AppendLine(@"				</tr>");

			// 'white groups one' groups.
			if (userValidator.WhiteGroupsOne.Length > 0) {
				html.AppendLine(@"				<tr>");
				html.AppendLine(@"					<th>White groups (one)</th>");
				html.AppendLine(@"					<td>");
				foreach (GroupPrincipal group in userValidator.WhiteGroupsOne) {
					html.AppendLine($"						{group.Name}<br>");
				}
				html.AppendLine($"						The action is only performed, if the user is member of one of the 'WhiteGroupsOne' groups");
				html.AppendLine(@"					</td>");
				html.AppendLine(@"				</tr>");
			}

			// 'white groups all' groups.
			if (userValidator.WhiteGroupsAll.Length > 0) {
				html.AppendLine(@"				<tr>");
				html.AppendLine(@"					<th>White groups (all)</th>");
				html.AppendLine(@"					<td>");
				foreach (GroupPrincipal group in userValidator.WhiteGroupsAll) {
					html.AppendLine($"						{group.Name}<br>");
				}
				html.AppendLine($"						The action is only performed, if the user is member of all of the 'WhiteGroupsAll' groups");
				html.AppendLine(@"					</td>");
				html.AppendLine(@"				</tr>");
			}

			// 'black groups one' groups.
			if (userValidator.BlackGroupsOne.Length > 0) {
				html.AppendLine(@"				<tr>");
				html.AppendLine(@"					<th>Black groups (one)</th>");
				html.AppendLine(@"					<td>");
				foreach (GroupPrincipal group in userValidator.BlackGroupsOne) {
					html.AppendLine($"						{group.Name}<br>");
				}
				html.AppendLine($"						The action is only performed, if the user is not member of one of the 'BlackGroupsOne' groups");
				html.AppendLine(@"					</td>");
				html.AppendLine(@"				</tr>");
			}

			// 'black groups all' groups.
			if (userValidator.BlackGroupsAll.Length > 0) {
				html.AppendLine(@"				<tr>");
				html.AppendLine(@"					<th>Black groups (all)</th>");
				html.AppendLine(@"					<td>");
				foreach (GroupPrincipal group in userValidator.BlackGroupsAll) {
					html.AppendLine($"						{group.Name}<br>");
				}
				html.AppendLine($"						The action is only performed, if the user is not member of any of the 'BlackGroupsAll' groups");
				html.AppendLine(@"					</td>");
				html.AppendLine(@"				</tr>");
			}

			// Base DN.
			if (configBaseDN.Length > 0) {
				html.AppendLine(@"				<tr>");
				html.AppendLine(@"					<th>Below DN</th>");
				html.AppendLine(@"					<td>");
				html.AppendLine($"						{configBaseDN}<br>");
				html.AppendLine(@"					</td>");
				html.AppendLine(@"				</tr>");
			}

			html.AppendLine(@"			</tbody>");
			html.AppendLine(@"		</table>");

			// Errors.
			if (errors.Count > 0) {
				html.AppendLine(@"		<h2>Errors</h2>");
				html.AppendLine(@"		<table border=""0"" cellpadding=""5"" cellspacing=""0"">");
				html.AppendLine(@"			<tbody>");
				foreach (String error in errors) {
					html.AppendLine(@"				<tr>");
					html.AppendLine(@"					<td>");
					html.AppendLine($"						{error}<br>");
					html.AppendLine(@"					</td>");
					html.AppendLine(@"				</tr>");
				}
				html.AppendLine(@"			</tbody>");
				html.AppendLine(@"		</table>");
			}

			html.AppendLine(@"	</body> ");
			html.AppendLine(@"</html> ");

			// Return the html.
			return html.ToString();
		} // GetHtmlMessage
		#endregion

	} // ActiveDirectoryInactiveUsers
	#endregion

} // NDK.PluginCollection