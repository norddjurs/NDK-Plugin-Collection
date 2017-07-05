using System;
using System.Collections.Generic;
using System.Text;
using NDK.Framework;
using NDK.Framework.CprBroker;

namespace NDK.PluginCollection {

	#region SendWindowsMessage class.
	public class CprSynchronize : BasePlugin {

		#region Implement PluginBase abstraction.
		/// <summary>
		/// Gets the unique plugin guid.
		/// When implementing a plugin, this method should return the same unique guid every time. 
		/// </summary>
		/// <returns></returns>
		public override Guid GetGuid() {
			return new Guid("{2525B770-5A44-40BD-B480-B0626C410684}");
		} // GetGuid

		/// <summary>
		/// Gets the the plugin name.
		/// When implementing a plugin, this method should return a proper display name.
		/// </summary>
		/// <returns></returns>
		public override String GetName() {
			return "NDK CPR Synchronize";
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
				// Get config.
				Boolean configMessageSend = this.GetLocalValue("MessageSend", true);
				List<String> configMessageTo = this.GetLocalValues("MessageTo");
				String configMessageSubject = this.GetLocalValue("MessageSubject", this.GetName());

				Boolean configFailAlways = this.GetLocalValue("FailAlways", true);
				String configInfoText = this.GetLocalValue("InfoText", "User automatically updated.");

				Boolean configSyncAd = this.GetLocalValue("SyncAD", false);
				Boolean configSyncAdOnlyEnabledUsers = this.GetLocalValue("SyncADOnlyEnabledUsers", true);
				Boolean configSyncAdFirstName = this.GetLocalValue("SyncADFirstName", true);
				Boolean configSyncAdLastName = this.GetLocalValue("SyncADLastName", true);
				Boolean configSyncAdFullName = this.GetLocalValue("SyncADFullName", true);
				Boolean configSyncAdDisplayName = this.GetLocalValue("SyncADDisplayName", true);
				String configSyncAdFilterBaseDN = this.GetLocalValue("SyncADFilterBaseDN", String.Empty);

				Boolean configSyncSofd = this.GetLocalValue("SyncSOFD", false);
				Boolean configSyncSofdOnlyEnabledUsers = this.GetLocalValue("SyncSOFDOnlyActiveEmployees", true);
				Boolean configSyncSofdFirstName = this.GetLocalValue("SyncSOFDFirstName", true);
				Boolean configSyncSofdLastName = this.GetLocalValue("SyncSOFDLastName", true);
				Boolean configSyncSofdFullName = this.GetLocalValue("SyncSOFDFullName", true);
				Boolean configSyncSofdDisplayName = this.GetLocalValue("SyncSOFDDisplayName", true);
				List<String> configSyncSofdFilterWhiteID = this.GetLocalValues("SyncSOFDFilterWhiteID");
				List<String> configSyncSofdFilterBlackID = this.GetLocalValues("SyncSOFDFilterBlackID");
				List<String> configSyncSofdFilterWhiteOrganization = this.GetLocalValues("SyncSOFDFilterWhiteOrganization");
				List<String> configSyncSofdFilterBlackOrganization = this.GetLocalValues("SyncSOFDFilterBlackOrganization");
				List<String> configSyncSofdFilterWhiteLeader = this.GetLocalValues("SyncSOFDFilterWhiteLeader");
				List<String> configSyncSofdFilterBlackLeader = this.GetLocalValues("SyncSOFDFilterBlackLeader");

				// Lowercase.
				for (Int32 index = 0; index < configSyncSofdFilterWhiteID.Count; index++) {
					configSyncSofdFilterWhiteID[index] = configSyncSofdFilterWhiteID[index].ToLower();
				}
				for (Int32 index = 0; index < configSyncSofdFilterBlackID.Count; index++) {
					configSyncSofdFilterBlackID[index] = configSyncSofdFilterBlackID[index].ToLower();
				}

				// Only synchronize users according to membership of the SyncWhiteGroupsOne, SyncWhiteGroupsAll, SyncBlackGroupsOne, SyncBlackGroupsAll groups.
				ActiveDirectoryUserValidator adUserValidator = new ActiveDirectoryUserValidator(this, "SyncADFilter");

				// Get options.
				DateTime optionSyncLastTime = this.GetOptionValue("SyncLastTime", DateTime.MinValue);

				// Report.
				HtmlBuilder html = new HtmlBuilder();
				List<List<String>> tableUpdated = new List<List<String>>();
				tableUpdated.Add(new List<String>() { "", "ID", "MA#", "Fields" });
				Int32 tableUpdatedAdCount = 0;
				Int32 tableUpdatedEmployeeCount = 0;


				//-----------------------------------------------------------------------------------------------------------------------------------
				// Synchronize users in Active Directory.
				//-----------------------------------------------------------------------------------------------------------------------------------
				if (configSyncAd == true) {
					foreach (AdUser adUser in this.GetAllUsers()) {
						if ((adUser.Modified.Value.CompareTo(optionSyncLastTime) > 0) &&																			// User updated after last synchronization.
							(this.GetUserCprNumber(adUser).IsNullOrWhiteSpace() == false) &&																		// User has a CPR number.
							(adUserValidator.ValidateUser(adUser) == true) &&																						// User membership of white/black groups.
							((configSyncAdFilterBaseDN.IsNullOrWhiteSpace() == true) || (adUser.DistinguishedName.EndsWith(configSyncAdFilterBaseDN) == true)) &&	// User is below base DN.
							((adUser.Enabled.Value == true) || (configSyncAdOnlyEnabledUsers == false))) {															// User enabled state.
							List<String> adUpdated = new List<String>();
							List<String> adUpdatedLong = new List<String>();

							// Query CPR register.
							CprSearchResult cpr = this.CprSearch(this.GetUserCprNumber(adUser));
							if (cpr != null) {
								if ((configSyncAdFirstName == true) && (adUser.GivenName.Equals((cpr.FirstName + " " + cpr.MiddleName).Trim()) == false)) {
									adUpdated.Add("First name");
									adUpdatedLong.Add(String.Format("First name ({0} -> {1})", adUser.GivenName, (cpr.FirstName + " " + cpr.MiddleName).Trim()));
									adUser.GivenName = (cpr.FirstName + " " + cpr.MiddleName).Trim();
								}

								if ((configSyncAdLastName == true) && (adUser.Surname.Equals(cpr.LastName) == false)) {
									adUpdated.Add("Last name");
									adUpdatedLong.Add(String.Format("Last name ({0} -> {1})", adUser.Surname, cpr.LastName));
									adUser.Surname = cpr.LastName;
								}

								if ((configSyncAdFullName == true) && (adUser.Name.Equals(cpr.FullName) == false)) {
									adUpdated.Add("Full name");
									adUpdatedLong.Add(String.Format("Full name ({0} -> {1})", adUser.Name, cpr.FullName));
									adUser.Name = cpr.FullName;
								}

								if ((configSyncAdDisplayName == true) && (adUser.DisplayName.Equals(cpr.FullName) == false)) {
									adUpdated.Add("Display name");
									adUpdatedLong.Add(String.Format("Display name ({0} -> {1})", adUser.DisplayName, cpr.FullName));
									adUser.DisplayName = cpr.FullName;
								}

								// Save.
								if (adUpdated.Count > 0) {
									if (configFailAlways == false) {
										// Update info text.
										if (configInfoText.IsNullOrWhiteSpace() == false) {
											adUser.InsertInfo(configInfoText + ": " + String.Join(", ", adUpdated));
										}

										// Save.
										//adUser.Save();

										// Log.
										this.LogDebug("Updated user in AD: {0} - {1}. Fields: {2}", adUser.SamAccountName, adUser.Name, String.Join(", ", adUpdated));
									}
								
									// Report.
									tableUpdated.Add(new List<String>() { "AD", adUser.SamAccountName, "-", String.Join(Environment.NewLine, adUpdatedLong) });
									tableUpdatedAdCount++;
								}
							}
						}
					}
				}


				//-----------------------------------------------------------------------------------------------------------------------------------
				// Synchronize users in SOFD Directory.
				//-----------------------------------------------------------------------------------------------------------------------------------
				if (configSyncSofd == true) {
					foreach (SofdEmployee employee in this.GetAllEmployees()) {
						if ((employee.SidstAendret.CompareTo(optionSyncLastTime) > 0) &&																// Employtt updated after last synchronization.
							(employee.CprNummer.IsNullOrWhiteSpace() == false) &&																		// Employee has a CPR number.
							((
								(configSyncSofdFilterWhiteID.Count == 0) &&																				// Empty identifier white list.
								(configSyncSofdFilterWhiteOrganization.Count == 0) &&																	// Empty organizatiion white list.
								(configSyncSofdFilterWhiteLeader.Count == 0)																			// Empty leader white list.
							) ||
								(configSyncSofdFilterWhiteID.Contains(employee.MedarbejderHistorikId.ToString()) == true) ||							// Identifier in white list.
								(configSyncSofdFilterWhiteID.Contains(employee.AdBrugerNavn.GetNotNull().ToLower()) == true) ||							// Identifier in white list.
								(configSyncSofdFilterWhiteID.Contains(employee.CprNummer.GetNotNull().FormatStringCpr()) == true) ||					// Identifier in white list.
								(configSyncSofdFilterWhiteID.Contains(employee.Epost.GetNotNull().ToLower()) == true) ||								// Identifier in white list.
								(configSyncSofdFilterWhiteID.Contains(employee.Uuid.ToString().ToLower()) == true) || 									// Identifier in white list.
								(configSyncSofdFilterWhiteOrganization.Contains(employee.OrganisationId.ToString()) == true) ||							// Organization identifier in white list.
								(configSyncSofdFilterWhiteOrganization.Contains(employee.OrganisationKortNavn.GetNotNull()) == true) ||					// Organization short name in white list.
								(configSyncSofdFilterWhiteOrganization.Contains(employee.OrganisationNavn.GetNotNull()) == true) || 					// Organization name in white list.
								(configSyncSofdFilterWhiteLeader.Contains(employee.NaermesteLederCprNummer.GetNotNull().FormatStringCpr()) == true) ||	// Leader CPR number in white list.
								(configSyncSofdFilterWhiteLeader.Contains(employee.NaermesteLederAdBrugerNavn.GetNotNull()) == true) ||					// Leader AD username in white list.
								(configSyncSofdFilterWhiteLeader.Contains(employee.NaermesteLederNavn.GetNotNull()) == true) 							// Leader name in white list.
							) &&
							(configSyncSofdFilterBlackID.Contains(employee.MedarbejderHistorikId.ToString()) == false) &&								// Identifier not in black list.
							(configSyncSofdFilterBlackID.Contains(employee.AdBrugerNavn.GetNotNull().ToLower()) == false) &&							// Identifier not in black list.
							(configSyncSofdFilterBlackID.Contains(employee.CprNummer.GetNotNull().FormatStringCpr()) == false) &&						// Identifier not in black list.
							(configSyncSofdFilterBlackID.Contains(employee.Epost.GetNotNull().ToLower()) == false) &&									// Identifier not in black list.
							(configSyncSofdFilterBlackID.Contains(employee.Uuid.ToString().ToLower()) == false) &&										// Identifier not in black list.
							(configSyncSofdFilterBlackOrganization.Contains(employee.OrganisationId.ToString()) == false) &&							// Organization identifier not in black list.
							(configSyncSofdFilterBlackOrganization.Contains(employee.OrganisationKortNavn.GetNotNull()) == false) &&					// Organization short name not in black list.
							(configSyncSofdFilterBlackOrganization.Contains(employee.OrganisationNavn.GetNotNull()) == false) &&						// Organization name not in black list.
							(configSyncSofdFilterBlackID.Contains(employee.NaermesteLederCprNummer.GetNotNull().FormatStringCpr()) == false) &&			// Leader CPR number not in black list.
							(configSyncSofdFilterBlackID.Contains(employee.NaermesteLederAdBrugerNavn.GetNotNull()) == false) &&						// Leader AD username not in black list.
							(configSyncSofdFilterBlackID.Contains(employee.NaermesteLederNavn.GetNotNull()) == false) &&								// Leader name not in black list.
							((employee.Aktiv == true) || (configSyncSofdOnlyEnabledUsers == false))) {													// Employee active state.
							List<String> employeeUpdated = new List<String>();
							List<String> employeeUpdatedLong = new List<String>();

							// Query CPR register.
							CprSearchResult cpr = this.CprSearch(employee.CprNummer);
							if (cpr != null) {
								if ((configSyncSofdFirstName == true) && (employee.ForNavn.GetNotNull().Equals((cpr.FirstName + " " + cpr.MiddleName).Trim()) == false)) {
									employeeUpdated.Add("First name");
									employeeUpdatedLong.Add(String.Format("First name ({0} -> {1})", employee.ForNavn, (cpr.FirstName + " " + cpr.MiddleName).Trim()));
									employee.ForNavn = (cpr.FirstName + " " + cpr.MiddleName).Trim();
								}

								if ((configSyncSofdLastName == true) && (employee.EfterNavn.GetNotNull().Equals(cpr.LastName) == false)) {
									employeeUpdated.Add("Last name");
									employeeUpdatedLong.Add(String.Format("Last name ({0} -> {1})", employee.EfterNavn, cpr.LastName));
									employee.EfterNavn = cpr.LastName;
								}

								if ((configSyncSofdFullName == true) && (employee.Navn.GetNotNull().Equals(cpr.FullName) == false)) {
									employeeUpdated.Add("Full name");
									employeeUpdatedLong.Add(String.Format("Full name ({0} -> {1})", employee.Navn, cpr.FullName));
									employee.Navn = cpr.FullName;
								}

								if ((configSyncSofdDisplayName == true) && (employee.KaldeNavn.GetNotNull().Equals(cpr.FullName) == false)) {
									employeeUpdated.Add("Display name");
									employeeUpdatedLong.Add(String.Format("Display name ({0} -> {1})", employee.KaldeNavn, cpr.FullName));
									employee.KaldeNavn = cpr.FullName;
								}

								// Save.
								if (employeeUpdated.Count > 0) {
									if (configFailAlways == false) {
										// Save.
										//employee.Save(true);

										// Log.
										this.LogDebug("Updated user in SOFD: {0} - {1}. Fields: {2}", employee.MedarbejderHistorikId, employee.Navn, String.Join(", ", employeeUpdated));
									}

									// Report.
									tableUpdated.Add(new List<String>() { "SOFD", employee.AdBrugerNavn, employee.MaNummer.ToString(), String.Join(Environment.NewLine, employeeUpdatedLong) });
									tableUpdatedEmployeeCount++;
								}
							}
						}
					}
				}


				// Report.
				html.AppendHeading2("Updated users/employees");
				tableUpdated.Add(new List<String>() { tableUpdatedAdCount + " AD users, " + tableUpdatedEmployeeCount + " SOFD users" });
				html.AppendHorizontalTable(tableUpdated, 1, 0);

				// Configuration.
				List<List<String>> tableConfig = new List<List<String>>();
				if (configFailAlways == true) {
					tableConfig.Add(new List<String>() { "Fail", "Do not write changes to the Active Directory and the SOFD Directory." });
				} else {
					tableConfig.Add(new List<String>() { "Fail", "Write changes to the Active Directory and the SOFD Directory." });
				}
				tableConfig.Add(new List<String>() { "Info text", configInfoText });

				tableConfig.Add(new List<String>() { "SyncAD", 	configSyncAd.ToString() });
				if (configSyncAd == true) {
					tableConfig.Add(new List<String>() { "SyncADOnlyEnabledUsers", configSyncAdOnlyEnabledUsers.ToString() });
					tableConfig.Add(new List<String>() { "SyncADFirstName", configSyncAdFirstName.ToString() });
					tableConfig.Add(new List<String>() { "SyncADLastName", configSyncAdLastName.ToString() });
					tableConfig.Add(new List<String>() { "SyncADFullName", configSyncAdFullName.ToString() });
					tableConfig.Add(new List<String>() { "SyncADDisplayName", configSyncAdDisplayName.ToString() });
					adUserValidator.AddOptionsToHtmlBuilderVerticalTable(tableConfig);
					tableConfig.Add(new List<String>() { "SyncADFilterBaseDN", configSyncAdFilterBaseDN });
				}

				tableConfig.Add(new List<String>() { "SyncSOFD", 	configSyncSofd.ToString() });
				if (configSyncSofd == true) {
					tableConfig.Add(new List<String>() { "SyncSOFDOnlyEnabledUsers", configSyncSofdOnlyEnabledUsers.ToString() });
					tableConfig.Add(new List<String>() { "SyncSOFDFirstName", configSyncSofdFirstName.ToString() });
					tableConfig.Add(new List<String>() { "SyncSOFDLastName", configSyncSofdLastName.ToString() });
					tableConfig.Add(new List<String>() { "SyncSOFDFullName", configSyncSofdFullName.ToString() });
					tableConfig.Add(new List<String>() { "SyncSOFDDisplayName", configSyncSofdDisplayName.ToString() });
				}

				html.AppendHeading2("Configuration");
				html.AppendVerticalTable(tableConfig);

				// Send message.
				if (configMessageSend == true) {
					if (configMessageTo.Count > 0) {
						this.SendMail(
							String.Join(";", configMessageTo.ToArray()),
							configMessageSubject,
							html.ToString(),
							true
						);
					} else {
						this.SendMail(
							configMessageSubject,
							html.ToString(),
							true
						);
					}
				}

				// Set options.
				this.SetOptionValue("SyncLastTime", DateTime.Now.Date);
			} catch (Exception exception) {
				// Send message on error.
				this.SendMail("Error " + this.GetName(), exception.Message, false);

				// Throw the error.
				throw;
			}
		} // Run
		#endregion

	} // CprSynchronize
	#endregion

} // NDK.PluginCollection