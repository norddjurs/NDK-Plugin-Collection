using System;
using NDK.Framework;
using System.Data;
using System.DirectoryServices.AccountManagement;
using System.Collections.Generic;
using System.Text;

namespace NDK.PluginCollection {

	#region DemoPlugin class.
	public class ActiveDirectoryInactiveUsers : PluginBase {

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
				List<String> optionMessageTo = this.GetConfigValues("MessageTo");
				String optionMessageSubject = this.GetConfigValue("MessageSubject", this.GetName());
				Boolean optionFailOnGroupNotFound = this.GetConfigValue("FailOnGroupNotFound", true);
				Boolean optionFailAlways = this.GetConfigValue("FailAlways", true);
				String optionBaseDN = this.GetConfigValue("BaseDN", String.Empty).Trim();
				String optionInfoText = this.GetConfigValue("InfoText", "User automatically disabled.").Trim();
				List <String> optionMemberOfGroupStrings = this.GetConfigValues("MemberOfGroup");
				List<String> optionNotMemberOfGroupStrings = this.GetConfigValues("NotMemberOfGroup");
				Int32 optionInactivePeriodDays = this.GetConfigValue("InactivePeriodDays", 90);
				String optionInactiveAction = this.GetConfigValue("InactiveAction", "DISABLE").ToLower();

				List<String> errors = new List<String>();

				foreach (GroupPrincipal group in this.GetAllGroups()) {
					this.Log("Group: {0}.", group.Name);
				}


				// Get the member of groups.
				List<GroupPrincipal> optionMemberOfGroups = new List<GroupPrincipal>();
				foreach (String optionMemberOfGroupString in optionMemberOfGroupStrings) {
					GroupPrincipal optionMemberOfGroup = this.GetGroup(optionMemberOfGroupString);
					if (optionMemberOfGroup != null) {
						// The group was found.
						optionMemberOfGroups.Add(optionMemberOfGroup);

						// Log.
						this.Log("User must be member of group '{0}'.", optionMemberOfGroup.Name);
					} else {
						// Log that the group was not found.
						this.LogError("Unable to find group '{0}'.", optionMemberOfGroupString);
					}
				}
				
				// Get the not member of groups.
				List <GroupPrincipal> optionNotMemberOfGroups = new List<GroupPrincipal>();
				foreach (String optionNotMemberOfGroupString in optionNotMemberOfGroupStrings) {
					GroupPrincipal optionNotMemberOfGroup = this.GetGroup(optionNotMemberOfGroupString);
					if (optionNotMemberOfGroup != null) {
						// The group was found.
						optionNotMemberOfGroups.Add(optionNotMemberOfGroup);

						// Log.
						this.Log("User may not be member of group '{0}'.", optionNotMemberOfGroup.Name);
					} else {
						// Log that the group was not found.
						this.LogError("Unable to find group '{0}'.", optionNotMemberOfGroupString);
					}
				}

				// Fail if one or more groups was not found.
				if ((optionFailOnGroupNotFound == true) &&
					((optionMemberOfGroups.Count != optionMemberOfGroupStrings.Count) || (optionNotMemberOfGroups.Count != optionNotMemberOfGroupStrings.Count))) {
					throw new Exception("One or more groups was not found in Active Directory.");
				}

				// Get the inactive date.
				if (optionInactivePeriodDays < 1) {
					optionInactivePeriodDays = 90;
				}
				DateTime optionInactivePeriod = DateTime.Now.AddDays(-optionInactivePeriodDays);

				// Log.
				this.Log("Processing inactive users for {0} days, since {1:yyyy-MM-dd}.", optionInactivePeriodDays, optionInactivePeriod);

				// Get all inactive users.
				Boolean userMemberOfGroups = true;
				Boolean userNotMemberOfGroups = true;
				Boolean userInBaseDN = true;
				List<Person> inactiveUsers = new List<Person>();
				foreach (Person user in this.GetAllUsers(UserQuery.ALL)) {
					if ((user.LastLogon != null) && (user.LastLogon.Value.CompareTo(optionInactivePeriod) < 0)) {
						userMemberOfGroups = true;
						userNotMemberOfGroups = true;
						userInBaseDN = true;

						// Log.
						this.Log("Inactive user: {0} - {1} ({2:yyyy-MM-dd})", user.SamAccountName, user.Name, user.LastLogon);

						// Validate that the user is member of one of the groups.
						foreach (GroupPrincipal group in optionMemberOfGroups) {
							if (this.IsUserMemberOfGroup(user, group, true) == false) {
								// Fail.
								userMemberOfGroups = false;

								// Log.
								this.Log("The user should be member of '{0}', but is not.", group.Name);
							}
						}

						// Validate that the user is not member of one of the groups.
						foreach (GroupPrincipal group in optionNotMemberOfGroups) {
							if (this.IsUserMemberOfGroup(user, group, true) == true) {
								// Fail.
								userNotMemberOfGroups = false;

								// Log.
								this.Log("The user should not be member of '{0}', but is.", group.Name);
							}
						}

						// Validate that the user is below the base DN.
						if ((optionBaseDN.Length > 0) && (user.DistinguishedName.EndsWith(optionBaseDN) == false)) {
							// Fail.
							userInBaseDN = false;

							// Log.
							this.Log("The user is not below the base DN '{0}'. User DN is '{1}'.", optionBaseDN, user.DistinguishedName);
						}

						// Add the user.
						if ((userMemberOfGroups == true) && (userNotMemberOfGroups == true) && (userInBaseDN == true)) {
							inactiveUsers.Add(user);
						}
					}
				}

				// Process found inactive users.
				if (optionFailAlways == false) {
					foreach (Person user in inactiveUsers) {
						switch (optionInactiveAction) {
							case "disable":
									// Log.
									this.Log("Disabling inactive user: {0} - {1} ({2:yyyy-MM-dd})", user.SamAccountName, user.Name, user.LastLogon);

								// Disable the user.
								try {
									//user.Enabled = false;
									//user.Info = String.Format("{0:yyyy-MM-dd} {0}", DateTime.Now, optionInfoText) + Environment.NewLine + user.Info;
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
									//user.Info = String.Format("{0:yyyy-MM-dd} {0}", DateTime.Now, optionInfoText) + Environment.NewLine + user.Info;
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
				if (optionMessageTo.Count > 0) {
					this.SendMail(
						String.Join(";", optionMessageTo.ToArray()),
						optionMessageSubject,
						this.GetHtmlMessage(optionFailAlways, optionBaseDN, optionMemberOfGroups, optionNotMemberOfGroups, optionInactivePeriodDays, optionInactiveAction, inactiveUsers, errors),
						true
					);
				} else {
					this.SendMail(
						optionMessageSubject,
						this.GetHtmlMessage(optionFailAlways, optionBaseDN, optionMemberOfGroups, optionNotMemberOfGroups, optionInactivePeriodDays, optionInactiveAction, inactiveUsers, errors),
						true
					);
				}

			} catch (Exception exception) {
				// Send message on error.
				this.SendMail("Error " + this.GetName(), exception.Message, false);

				// Throw the error.
				throw;
			}
		} // Run

		private String GetHtmlMessage(Boolean optionFailAlways, String optionBaseDN, List<GroupPrincipal> optionMemberOfGroups, List<GroupPrincipal> optionNotMemberOfGroups, Int32 optionInactivePeriodDays, String optionInactiveAction, List<Person> inactiveUsers, List<String> errors) {
			DateTime optionInactivePeriod = DateTime.Now.AddDays(-optionInactivePeriodDays);
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
			if (optionFailAlways == false) {
				html.AppendLine(@"		<p>This automatic task has performed the specified action on the inactive users in the active directory.");
				html.AppendLine(@"		</p>");
			} else {
				html.AppendLine(@"		<p>This automatic task is DEACTIVATED.<br>");
				html.AppendLine(@"		<p>When it is enabled, it will performed the specified action on the inactive users in the active directory.");
				html.AppendLine(@"		</p>");
			}

			// Inactive users.
			html.AppendLine(@"		<h2>Inactive users</h2>");
			html.AppendLine(@"		<table border=""0"" cellpadding=""10"" cellspacing=""0"">");
			html.AppendLine(@"			<thead>");
			html.AppendLine(@"				<tr>");
			html.AppendLine(@"					<th>");
			html.AppendLine($"						Userid<br>");
			html.AppendLine(@"					</th>");

			html.AppendLine(@"					<th>");
			html.AppendLine($"						Full name<br>");
			html.AppendLine(@"					</th>");

			html.AppendLine(@"					<th>");
			html.AppendLine($"						Last logon<br>");
			html.AppendLine(@"					</th>");
			html.AppendLine(@"				</tr>");
			html.AppendLine(@"			</thead>");
			html.AppendLine(@"			<tbody>");

			foreach (Person user in inactiveUsers) {
				html.AppendLine(@"				<tr>");
				html.AppendLine(@"					<td>");
				html.AppendLine($"						{user.SamAccountName}<br>");
				html.AppendLine(@"					</td>");

				html.AppendLine(@"					<td>");
				html.AppendLine($"						{user.Name}<br>");
				html.AppendLine(@"					</td>");

				html.AppendLine(@"					<td>");
				html.AppendLine($"						{user.LastLogon.Value:yyyy-MM-dd}<br>");
				html.AppendLine(@"					</td>");
				html.AppendLine(@"				</tr>");
			}
			html.AppendLine(@"			</tbody>");
			html.AppendLine(@"		</table>");

			html.AppendLine(@"		<h2>Configuration</h2>");
			html.AppendLine(@"		<table border=""0"" cellpadding=""10"" cellspacing=""0"">");
			html.AppendLine(@"			<tbody>");

			// Period.
			html.AppendLine(@"				<tr>");
			html.AppendLine(@"					<th>Inactive period</th>");
			html.AppendLine(@"					<td>");
			html.AppendLine($"						{optionInactivePeriodDays} days, since {optionInactivePeriod:yyyy-MM-dd}<br>");
			html.AppendLine(@"					</td>");
			html.AppendLine(@"				</tr>");

			// Action.
			html.AppendLine(@"				<tr>");
			html.AppendLine(@"					<th>Inactive action</th>");
			html.AppendLine(@"					<td>");
			html.AppendLine($"						{optionInactiveAction.ToUpper()} the inactive users<br>");
			html.AppendLine(@"					</td>");
			html.AppendLine(@"				</tr>");

			// Member of group.
			if (optionMemberOfGroups.Count > 0) {
				html.AppendLine(@"				<tr>");
				html.AppendLine(@"					<th>Member of group</th>");
				html.AppendLine(@"					<td>");
				foreach (GroupPrincipal group in optionMemberOfGroups) {
					html.AppendLine($"						{group.Name}<br>");
				}
				html.AppendLine(@"					</td>");
				html.AppendLine(@"				</tr>");
			}

			// Not member of group.
			if (optionNotMemberOfGroups.Count > 0) {
				html.AppendLine(@"				<tr>");
				html.AppendLine(@"					<th>Not member of group</th>");
				html.AppendLine(@"					<td>");
				foreach (GroupPrincipal group in optionNotMemberOfGroups) {
					html.AppendLine($"						{group.Name}<br>");
				}
				html.AppendLine(@"					</td>");
				html.AppendLine(@"				</tr>");
			}

			// Base DN.
			if (optionBaseDN.Length > 0) {
				html.AppendLine(@"				<tr>");
				html.AppendLine(@"					<th>Below DN</th>");
				html.AppendLine(@"					<td>");
				html.AppendLine($"						{optionBaseDN}<br>");
				html.AppendLine(@"					</td>");
				html.AppendLine(@"				</tr>");
			}

			html.AppendLine(@"			</tbody>");
			html.AppendLine(@"		</table>");

			// Errors.
			if (errors.Count > 0) {
				html.AppendLine(@"		<h2>Errors</h2>");
				html.AppendLine(@"		<table border=""0"" cellpadding=""10"" cellspacing=""0"">");
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