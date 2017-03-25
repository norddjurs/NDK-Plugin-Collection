using System;
using NDK.Framework;
using System.Data;
using System.DirectoryServices.AccountManagement;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace NDK.PluginCollection {

	#region ActiveDirectoryInactiveUsers class.
	public class ActiveDirectorySynchronizeWithSofd : BasePlugin {

		#region Implement PluginBase abstraction.
		/// <summary>
		/// Gets the unique plugin guid.
		/// When implementing a plugin, this method should return the same unique guid every time. 
		/// </summary>
		/// <returns></returns>
		public override Guid GetGuid() {
			return new Guid("{A7461D24-EF8C-440D-A004-F25531DB9A57}");
		} // GetGuid

		/// <summary>
		/// Gets the the plugin name.
		/// When implementing a plugin, this method should return a proper display name.
		/// </summary>
		/// <returns></returns>
		public override String GetName() {
			return "NDK ActiveDirectory Synchronize With SOFD";
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
				String configInfoText = this.GetLocalValue("InfoText", "User automatically updated.").Trim();

				// Get the saved option values.
				DateTime optionLastChanged = this.GetOptionValue("LastChanged", DateTime.MinValue);
				DateTime optionNewLastChanged = optionLastChanged;


				// Synchronize all users that exist both in the Active Directory and in SOFD.
				List<SofdEmployee> employees = this.GetAllEmployees(
					new SofdEmployeeFilter_AdBrugerNavn(SqlWhereFilterOperator.AND, SqlWhereFilterValueOperator.NotEquals, String.Empty),
					new SofdEmployeeFilter_SidstAendret(SqlWhereFilterOperator.AND, SqlWhereFilterValueOperator.GreaterThan, optionLastChanged),
					new SofdEmployeeFilter_Aktiv(SqlWhereFilterOperator.AND, SqlWhereFilterValueOperator.Equals, true)
				);
				foreach (SofdEmployee employee in employees) {
					// Get the user.
					AdUser user = this.GetUser(employee.AdBrugerNavn);
					if (user == null) {
						// The user does not exist in the active directory.
						this.Log("User not found: {2:yyyy-MM-dd}  '{0} - {1}'.", employee.AdBrugerNavn, employee.Navn, employee.SidstAendret);
					} else {
						// The user was found in the active directory.
						this.Log("Synchronize user: {2:yyyy-MM-dd}  '{0} - {1}'.", employee.AdBrugerNavn, employee.Navn, employee.SidstAendret);




						// Update the new last changed date, so it becomes the highest 'LastChanged' date.
						if (optionNewLastChanged.CompareTo(employee.SidstAendret) < 0) {
							optionNewLastChanged = employee.SidstAendret;
						}
					}
				}



				// Save the option values.
				this.SetOptionValue("LastChanged", optionNewLastChanged);

			} catch (Exception exception) {
				// Send message on error.
				this.SendMail("Error " + this.GetName(), exception.Message, false);

				// Throw the error.
				throw;
			}
		} // Run

		private String GetHtmlMessage(Boolean configFailAlways, String configBaseDN, ActiveDirectoryUserValidator userValidator, List<AdUser> inactiveUsers, List<String> errors) {
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
				html.AppendLine(@"		<p>This automatic task has performed the specified action on the users in the active directory.");
				html.AppendLine(@"		</p>");
			} else {
				html.AppendLine(@"		<p>This automatic task is DEACTIVATED.<br>");
				html.AppendLine(@"		<p>When it is enabled, it will perform the specified action on the users in the active directory.");
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
			//			html.AppendLine(@"				<tr>");
			//			html.AppendLine(@"					<th>Inactive period</th>");
			//			html.AppendLine(@"					<td>");
			//			html.AppendLine($"						{configInactivePeriodDays} days, since {configInactivePeriod:yyyy-MM-dd}<br>");
			//			html.AppendLine(@"					</td>");
			//			html.AppendLine(@"				</tr>");

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

	} // ActiveDirectorySynchronizeWithSofd
	#endregion

} // NDK.PluginCollection