using System;
using System.Collections.Generic;
using NDK.Framework;

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
				// Get config.
				Boolean configMessageSend = this.GetLocalValue("MessageSend", true);
				List<String> configMessageTo = this.GetLocalValues("MessageTo");
				String configMessageSubject = this.GetLocalValue("MessageSubject", this.GetName());

				Boolean configFailAlways = this.GetLocalValue("FailAlways", true);
				String configInfoText = this.GetLocalValue("InfoText", "User automatically updated.");

				List<String> configSyncFilterOrganizationWhite = this.GetLocalValues("SyncFilterOrganizationWhite");
				List<String> configSyncFilterOrganizationBlack = this.GetLocalValues("SyncFilterOrganizationBlack");
				List<String> configSyncFilterOmkostningsstedWhite = this.GetLocalValues("SyncFilterOmkostningsstedWhite");
				List<String> configSyncFilterOmkostningsstedBlack = this.GetLocalValues("SyncFilterOmkostningsstedBlack");
				Int64 configSyncFilterOmkostningsstedMin = this.GetLocalValue("SyncFilterOmkostningsstedMin", Int64.MinValue);
				Int64 configSyncFilterOmkostningsstedMax = this.GetLocalValue("SyncFilterOmkostningsstedMax", Int64.MaxValue);

				String configSyncCprNumber = this.GetLocalValue("SyncCprNumber", String.Empty);
				String configSyncMiFareId = this.GetLocalValue("SyncMiFareId", String.Empty);
				String configSyncFirstName = this.GetLocalValue("SyncFirstName", String.Empty);
				String configSyncLastName = this.GetLocalValue("SyncLastName", String.Empty);
				String configSyncFullName = this.GetLocalValue("SyncFullName", String.Empty);
				String configSyncDisplayName = this.GetLocalValue("SyncDisplayName", String.Empty);
				String configSyncPhone = this.GetLocalValue("SyncPhone", String.Empty);
				String configSyncMobile1 = this.GetLocalValue("SyncMobile1", String.Empty);
				String configSyncMobile2 = this.GetLocalValue("SyncMobile2", String.Empty);
				String configSyncMail = this.GetLocalValue("SyncMail", String.Empty);
				String configSyncTitle = this.GetLocalValue("SyncTitle", String.Empty);
				String configSyncDepartment = this.GetLocalValue("SyncDepartment", String.Empty);
				String configSyncManager = this.GetLocalValue("SyncManager", String.Empty);
				String configSyncAddress = this.GetLocalValue("SyncAddress", String.Empty);


				//-----------------------------------------------------------------------------------------------------------------------------------
				// Synchronize users existing both in Active Directory and SOFD Directory.
				//-----------------------------------------------------------------------------------------------------------------------------------
				// Report.
				HtmlBuilder html = new HtmlBuilder();
				List<List<String>> tableUpdated = new List<List<String>>();
				tableUpdated.Add(new List<String>() { "", "ID", "MA#", "Omk.sted", "Full name", "Fields" });
				Int32 tableUpdatedAdCount = 0;
				Int32 tableUpdatedEmployeeCount = 0;

				// Only synchronize users according to membership of the SyncWhiteGroupsOne, SyncWhiteGroupsAll, SyncBlackGroupsOne, SyncBlackGroupsAll groups.
				ActiveDirectoryUserValidator adUserValidator = new ActiveDirectoryUserValidator(this, "SyncFilter");

				// Only synchronize users one time.
				List<String> adUserSynchronized = new List<String>();

				// Get all active employees from SOFD Directory.
				List<SqlWhereFilterBase> employeesWhereFilters = new List<SqlWhereFilterBase>();
				employeesWhereFilters.Add(new SofdEmployeeFilter_Aktiv(SqlWhereFilterOperator.AND, SqlWhereFilterValueOperator.Equals, true));
				employeesWhereFilters.Add(new SofdEmployeeFilter_AnsaettelseAktiv(SqlWhereFilterOperator.AND, SqlWhereFilterValueOperator.Equals, true));
				List<SofdEmployee> employees = this.GetAllEmployees(employeesWhereFilters.ToArray());

				foreach (SofdEmployee employee in employees) {
					// Get the associated user in Active Directory.
					SofdOrganization organization = employee.GetOrganisation();
					AdUser adUser = this.GetUser(employee.AdBrugerNavn);
					if ((organization != null) &&
						(adUser != null) &&
						(adUserSynchronized.Contains(adUser.SamAccountName) == false) &&
						// Filter on organnization "Omkostningssted" between minimum and maximum values, or one of the WHITE lists.
						(
							((organization.Omkostningssted >= configSyncFilterOmkostningsstedMin) && (organization.Omkostningssted <= configSyncFilterOmkostningsstedMax)) ||
							// Filter on organnization "Omkostningssted" white list.
							(configSyncFilterOmkostningsstedWhite.Contains(organization.Omkostningssted.ToString()) == true) ||
							// Filter on organization identifier and names white list.
							(configSyncFilterOrganizationWhite.Contains(organization.OrganisationId.ToString()) == true) ||
							(configSyncFilterOrganizationWhite.Contains(organization.Navn) == true) ||
							(configSyncFilterOrganizationWhite.Contains(organization.KortNavn) == true)
						) &&
						// Filter on organization identifier and names black list.
						(configSyncFilterOrganizationBlack.Contains(organization.OrganisationId.ToString()) == false) &&
						(configSyncFilterOrganizationBlack.Contains(organization.Navn) == false) &&
						(configSyncFilterOrganizationBlack.Contains(organization.KortNavn) == false) &&
						(configSyncFilterOmkostningsstedBlack.Contains(organization.Omkostningssted.ToString()) == false) &&
						// Filter on Active Directory groups white and black lists.
						(adUserValidator.ValidateUser(adUser) == true)) {
						List<String> adUpdated = new List<String>();
						List<String> adUpdatedLong = new List<String>();
						List<String> employeeUpdated = new List<String>();
						List<String> employeeUpdatedLong = new List<String>();

						// Only synchronize users one time.
						adUserSynchronized.Add(adUser.SamAccountName);

						// Update.
						if (this.GetUserCprNumber(adUser).GetNotNull().Trim().Replace("-", String.Empty).Equals(employee.CprNummer.GetNotNull().Trim().Replace("-", String.Empty)) == false) {
							if (configSyncCprNumber.ToLower().Equals("ad") == true) {
								adUpdated.Add("CPR number");
								adUpdatedLong.Add(String.Format("CPR number ({0} -> {1})", this.GetUserCprNumber(adUser), employee.CprNummer.FormatStringCpr()));
								this.SetUserCprNumber(adUser, employee.CprNummer);
							}
							if (configSyncCprNumber.ToLower().Equals("sofd") == true) {
								employeeUpdated.Add("CPR number");
								employeeUpdatedLong.Add(String.Format("CPR number ({0} -> {1})", employee.CprNummer.FormatStringCpr(), this.GetUserCprNumber(adUser)));
								employee.CprNummer = this.GetUserCprNumber(adUser);
							}
						}

						if (this.GetUserMiFareId(adUser).GetNotNull().Trim().Equals(employee.MiFareId.GetNotNull().Trim()) == false) {
							if (configSyncMiFareId.ToLower().Equals("ad") == true) {
								adUpdated.Add("MiFare identifier");
								adUpdatedLong.Add(String.Format("MiFare identifier ({0} -> {1})", this.GetUserMiFareId(adUser), employee.MiFareId));
								this.SetUserMiFareId(adUser, employee.MiFareId);
							}
							if (configSyncMiFareId.ToLower().Equals("sofd") == true) {
								employeeUpdated.Add("MiFare identifier");
								employeeUpdatedLong.Add(String.Format("MiFare identifier ({0} -> {1})", employee.MiFareId, this.GetUserMiFareId(adUser)));
								employee.MiFareId = this.GetUserMiFareId(adUser);
							}
						}

						if (adUser.GivenName.GetNotNull().Trim().Equals(employee.ForNavn.GetNotNull().Trim()) == false) {
							if (configSyncFirstName.ToLower().Equals("ad") == true) {
								adUpdated.Add("First name");
								adUpdatedLong.Add(String.Format("First name ({0} -> {1})", adUser.GivenName, employee.ForNavn));
								adUser.GivenName = employee.ForNavn;
							}
							if (configSyncFirstName.ToLower().Equals("sofd") == true) {
								employeeUpdated.Add("First name");
								employeeUpdatedLong.Add(String.Format("First name ({0} -> {1})", employee.ForNavn, adUser.GivenName));
								employee.ForNavn = adUser.GivenName;
							}
						}

						if (adUser.Surname.GetNotNull().Trim().Equals(employee.EfterNavn.GetNotNull().Trim()) == false) {
							if (configSyncLastName.ToLower().Equals("ad") == true) {
								adUpdated.Add("Last name");
								adUpdatedLong.Add(String.Format("Last name ({0} -> {1})", adUser.Surname, employee.EfterNavn));
								adUser.Surname = employee.EfterNavn;
							}
							if (configSyncLastName.ToLower().Equals("sofd") == true) {
								employeeUpdated.Add("Last name");
								employeeUpdatedLong.Add(String.Format("Last name ({0} -> {1})", employee.EfterNavn, adUser.Surname));
								employee.EfterNavn = adUser.Surname;
							}
						}

						if (adUser.Name.GetNotNull().Trim().Equals(employee.Navn.GetNotNull().Trim()) == false) {
							if (configSyncFullName.ToLower().Equals("ad") == true) {
								adUpdated.Add("Full name");
								adUpdatedLong.Add(String.Format("Full name ({0} -> {1})", adUser.Name, employee.Navn));
								adUser.Name = employee.Navn;
							}
							if (configSyncFullName.ToLower().Equals("sofd") == true) {
								employeeUpdated.Add("Full name");
								employeeUpdatedLong.Add(String.Format("Full name ({0} -> {1})", employee.Navn, adUser.Name));
								employee.Navn = adUser.Name;
							}
						}

						if (adUser.DisplayName.GetNotNull().Trim().Equals(employee.KaldeNavn.GetNotNull().Trim()) == false) {
							if (configSyncDisplayName.ToLower().Equals("ad") == true) {
								adUpdated.Add("Display name");
								adUpdatedLong.Add(String.Format("Display name ({0} -> {1})", adUser.DisplayName, employee.KaldeNavn));
								adUser.DisplayName = employee.KaldeNavn;
							}
							if (configSyncDisplayName.ToLower().Equals("sofd") == true) {
								employeeUpdated.Add("Display name");
								employeeUpdatedLong.Add(String.Format("Display name ({0} -> {1})", employee.KaldeNavn, adUser.DisplayName));
								employee.KaldeNavn = adUser.DisplayName;
							}
						}

						if (adUser.TelephoneNumber.GetNotNull().Trim().Replace(" ", String.Empty).Equals(employee.TelefonNummer.GetNotNull().Trim().Replace(" ", String.Empty)) == false) {
							if (configSyncPhone.ToLower().Equals("ad") == true) {
								adUpdated.Add("Phone");
								adUpdatedLong.Add(String.Format("Phone ({0} -> {1})", adUser.TelephoneNumber.FormatStringPhone(), employee.TelefonNummer.FormatStringPhone()));
								adUser.TelephoneNumber = employee.TelefonNummer.FormatStringPhone();
							}
							if (configSyncPhone.ToLower().Equals("sofd") == true) {
								employeeUpdated.Add("Phone");
								employeeUpdatedLong.Add(String.Format("Phone ({0} -> {1})", employee.TelefonNummer.FormatStringPhone(), adUser.TelephoneNumber.FormatStringPhone()));
								employee.TelefonNummer = adUser.TelephoneNumber.FormatStringPhone();
							}
						}

						if (adUser.Mobile.GetNotNull().Trim().Replace(" ", String.Empty).Equals(employee.MobilNummer.GetNotNull().Trim().Replace(" ", String.Empty)) == false) {
							if (configSyncMobile1.ToLower().Equals("ad") == true) {
								adUpdated.Add("Mobile 1");
								adUpdatedLong.Add(String.Format("Mobile 1 ({0} -> {1})", adUser.Mobile.FormatStringPhone(), employee.MobilNummer.FormatStringPhone()));
								adUser.Mobile = employee.MobilNummer.FormatStringPhone();
							}
							if (configSyncMobile1.ToLower().Equals("sofd") == true) {
								employeeUpdated.Add("Mobile 1");
								employeeUpdatedLong.Add(String.Format("Mobile 1 ({0} -> {1})", employee.MobilNummer.FormatStringPhone(), adUser.Mobile.FormatStringPhone()));
								employee.MobilNummer = adUser.Mobile.FormatStringPhone();
							}
						}

						if (this.GetUserOtherMobile(adUser).Trim().Replace(" ", String.Empty).Equals(employee.MobilNummer2.GetNotNull().Trim().Replace(" ", String.Empty)) == false) {
							if (configSyncMobile2.ToLower().Equals("ad") == true) {
								adUpdated.Add("Mobile 2");
								adUpdatedLong.Add(String.Format("Mobile 2 ({0} -> {1})", this.GetUserOtherMobile(adUser), employee.MobilNummer2.FormatStringPhone()));
								this.SetUserOtherMobile(adUser, employee.MobilNummer2.FormatStringPhone());
							}
							if (configSyncMobile2.ToLower().Equals("sofd") == true) {
								employeeUpdated.Add("Mobile 2");
								employeeUpdatedLong.Add(String.Format("Mobile 2 ({0} -> {1})", employee.MobilNummer2.FormatStringPhone(), this.GetUserOtherMobile(adUser)));
								employee.MobilNummer2 = this.GetUserOtherMobile(adUser);
							}
						}

						if (adUser.SmtpProxyAddress.GetNotNull().Trim().ToUpper().Equals(employee.Epost.GetNotNull().Trim().ToUpper()) == false) {
							if (configSyncMail.ToLower().Equals("ad") == true) {
								adUpdated.Add("Mail");
								adUpdatedLong.Add(String.Format("Mail ({0} -> {1})", adUser.SmtpProxyAddress, employee.Epost.GetNotNull().ToUpper()));
								adUser.SmtpProxyAddress = employee.Epost.GetNotNull().ToUpper();
							}
							if (configSyncMail.ToLower().Equals("sofd") == true) {
								employeeUpdated.Add("Mail");
								employeeUpdatedLong.Add(String.Format("Mail ({0} -> {1})", employee.Epost.GetNotNull(), adUser.SmtpProxyAddress.GetNotNull().ToUpper()));
								employee.Epost = adUser.SmtpProxyAddress.GetNotNull().ToUpper();
							}
						}

						if (adUser.Title.GetNotNull().Trim().Equals(employee.StillingsBetegnelse.GetNotNull().Trim()) == false) {
							if (configSyncTitle.ToLower().Equals("ad") == true) {
								adUpdated.Add("Title");
								adUpdatedLong.Add(String.Format("Title ({0} -> {1})", adUser.Title, employee.StillingsBetegnelse));
								adUser.Title = employee.StillingsBetegnelse;
							}
							if (configSyncTitle.ToLower().Equals("sofd") == true) {
								employeeUpdated.Add("Title");
								employeeUpdatedLong.Add(String.Format("Title ({0} -> {1})", employee.StillingsBetegnelse, adUser.Title));
								employee.StillingsBetegnelse = adUser.Title;
							}
						}

						if (adUser.Department.GetNotNull().Trim().Equals(employee.OrganisationNavn.GetNotNull().Trim()) == false) {
							if (configSyncDepartment.ToLower().Equals("ad") == true) {
								adUpdated.Add("Department");
								adUpdatedLong.Add(String.Format("Department ({0} -> {1})", adUser.Department, employee.OrganisationNavn));
								adUser.Department = employee.OrganisationNavn;
							}
							if (configSyncDepartment.ToLower().Equals("sofd") == true) {
								employeeUpdated.Add("Department");
								employeeUpdatedLong.Add(String.Format("Department ({0} -> {1})", employee.OrganisationNavn, adUser.Department));
								employee.OrganisationNavn = adUser.Department;
							}
						}

						if ((configSyncManager.ToLower().Equals("ad") == true) || (configSyncManager.ToLower().Equals("sofd") == true)) {
							String adUserManagerDn = adUser.Manager.GetNotNull();
							String employeeManagerDn = String.Empty;
							try {
								employeeManagerDn = this.GetUser(employee.GetNearestLeader().AdBrugerNavn).DistinguishedName.GetNotNull();
							} catch {}

							if (adUserManagerDn.Equals(employeeManagerDn) == false) {
								if (configSyncManager.ToLower().Equals("ad") == true) {
									adUpdated.Add("Manager");
									adUpdatedLong.Add(String.Format("Manager ({0} -> {1})", adUserManagerDn, employeeManagerDn));
									adUser.Manager = employeeManagerDn;
								}
								if (configSyncManager.ToLower().Equals("sofd") == true) {
									AdUser adUserManager = this.GetUser(adUserManagerDn);
									if (adUserManager != null) {
										SofdEmployee employeeManager = this.GetEmployee(adUserManager.SamAccountName);

										employeeUpdated.Add("Manager");
										employeeUpdatedLong.Add(String.Format("Manager ({0} -> {1})", employeeManagerDn, adUserManagerDn));
										employee.NaermesteLederAdBrugerNavn = employeeManager.AdBrugerNavn;
										employee.NaermesteLederCprNummer = employeeManager.CprNummer;
										employee.NaermesteLederMaNummer = employeeManager.MaNummer;
										employee.NaermesteLederNavn = employeeManager.Navn;
									}
								}
							}
						}

						if (configSyncAddress.ToLower().Equals("work") == true) {
							if (organization != null) {
								String orgaizationAddress = (organization.Gade.GetNotNull() + Environment.NewLine + organization.StedNavn.GetNotNull()).Trim().Trim('\r', '\n').Trim();
								if (adUser.Street.GetNotNull().Trim().Equals(orgaizationAddress) == false) {
									adUpdated.Add("Street");
									adUpdatedLong.Add(String.Format("Street ({0} -> {1})", adUser.Street, orgaizationAddress));
									adUser.Street = orgaizationAddress;
								}
								if (adUser.City.GetNotNull().Trim().Equals(organization.By.GetNotNull().Trim()) == false) {
									adUpdated.Add("City");
									adUpdatedLong.Add(String.Format("City ({0} -> {1})", adUser.City, organization.By));
									adUser.City = organization.By;
								}
								if (adUser.PostalCode.GetNotNull().Trim().Equals(organization.PostNummer.ToString()) == false) {
									adUpdated.Add("Postal number");
									adUpdatedLong.Add(String.Format("Postal number ({0} -> {1})", adUser.PostalCode, organization.PostNummer));
									adUser.PostalCode = organization.PostNummer.ToString();
								}
								if (adUser.Country.GetNotNull().Trim().Equals("DK") == false) {	// Does not exist in SofdOrganization!
									adUpdated.Add("Country");
									adUpdatedLong.Add(String.Format("Country ({0} -> {1})", adUser.Country, "DK"));
									adUser.Country = "DK";
								}
							}
						}
						if (configSyncAddress.ToLower().Equals("home") == true) {
							String employeeAddress = (employee.Adresse.GetNotNull() + Environment.NewLine + employee.StedNavn.GetNotNull()).Trim().Trim('\r', '\n').Trim();
							if (adUser.Street.GetNotNull().Trim().Equals(employeeAddress) == false) {
								adUpdated.Add("Street");
								adUpdatedLong.Add(String.Format("Street ({0} -> {1})", adUser.Street, employeeAddress));
								adUser.Street = employeeAddress;
							}
							if (adUser.City.GetNotNull().Trim().Equals(employee.By.GetNotNull().Trim()) == false) {
								adUpdated.Add("City");
								adUpdatedLong.Add(String.Format("City ({0} -> {1})", adUser.City, employee.By));
								adUser.City = employee.By;
							}
							if (adUser.PostalCode.GetNotNull().Trim().Equals(employee.PostNummer.GetNotNull().Trim()) == false) {
								adUpdated.Add("Postal number");
								adUpdatedLong.Add(String.Format("Postal number ({0} -> {1})", adUser.PostalCode, employee.PostNummer));
								adUser.PostalCode = employee.PostNummer;
							}
							if (adUser.Country.GetNotNull().Trim().Equals(employee.Land.GetNotNull().Trim()) == false) {
								adUpdated.Add("Country");
								adUpdatedLong.Add(String.Format("Country ({0} -> {1})", adUser.Country, employee.Land));
								adUser.Country = employee.Land;
							}
						}

						// Save.
						if (adUpdated.Count > 0) {
							if (configFailAlways == false) {
								// Update info text.
								if (configInfoText.IsNullOrWhiteSpace() == false) {
									adUser.InsertInfo(configInfoText + ": " + String.Join(", ", adUpdated));
								}

								// Save.
								adUser.Save();

								// Log.
								this.LogDebug("Updated user in AD: {0} - {1}. Fields: {2}", adUser.SamAccountName, adUser.Name, String.Join(", ", adUpdated));
							}

							// Report.
							tableUpdated.Add(new List<String>() { "AD", adUser.SamAccountName, employee.MaNummer.ToString(), organization.Omkostningssted.ToString(), adUser.Name, String.Join(Environment.NewLine, adUpdatedLong) });
							tableUpdatedAdCount++;
						}
						if (employeeUpdated.Count > 0) {
							if (configFailAlways == false) {
								// Save.
								employee.Save(true);

								// Log.
								this.LogDebug("Updated user in SOFD: {0} - {1}. Fields: {2}", employee.MedarbejderHistorikId, employee.Navn, String.Join(", ", employeeUpdated));
							}

							// Report.
							tableUpdated.Add(new List<String>() { "SOFD", employee.AdBrugerNavn, employee.MaNummer.ToString(), organization.Omkostningssted.ToString(), employee.Navn, String.Join(Environment.NewLine, employeeUpdatedLong) });
							tableUpdatedEmployeeCount++;
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
				tableConfig.Add(new List<String>() { "SyncFilterOrganizationWhite", String.Join(Environment.NewLine, configSyncFilterOrganizationWhite) });
				tableConfig.Add(new List<String>() { "SyncFilterOrganizationBlack", String.Join(Environment.NewLine, configSyncFilterOrganizationBlack) });
				tableConfig.Add(new List<String>() { "SyncFilterOmkostningsstedWhite", String.Join(Environment.NewLine, configSyncFilterOmkostningsstedWhite) });
				tableConfig.Add(new List<String>() { "SyncFilterOmkostningsstedBlack", String.Join(Environment.NewLine, configSyncFilterOmkostningsstedBlack) });
				tableConfig.Add(new List<String>() { "SyncFilterOmkostningsstedMin", configSyncFilterOmkostningsstedMin.ToString() });
				tableConfig.Add(new List<String>() { "SyncFilterOmkostningsstedMax", configSyncFilterOmkostningsstedMax.ToString() });
				adUserValidator.AddOptionsToHtmlBuilderVerticalTable(tableConfig);
				tableConfig.Add(new List<String>() { "SyncCprNumber", configSyncCprNumber });
				tableConfig.Add(new List<String>() { "SyncMiFareId", configSyncMiFareId });
				tableConfig.Add(new List<String>() { "SyncFirstName", configSyncFirstName });
				tableConfig.Add(new List<String>() { "SyncLastName", configSyncLastName });
				tableConfig.Add(new List<String>() { "SyncFullName", configSyncFullName });
				tableConfig.Add(new List<String>() { "SyncDisplayName", configSyncDisplayName });
				tableConfig.Add(new List<String>() { "SyncPhone", configSyncPhone });
				tableConfig.Add(new List<String>() { "SyncMobile1", configSyncMobile1 });
				tableConfig.Add(new List<String>() { "SyncMobile2", configSyncMobile2 });
				tableConfig.Add(new List<String>() { "SyncMail", configSyncMail });
				tableConfig.Add(new List<String>() { "SyncTitle", configSyncTitle });
				tableConfig.Add(new List<String>() { "SyncDepartment", configSyncDepartment });
				tableConfig.Add(new List<String>() { "SyncManager", configSyncManager });
				tableConfig.Add(new List<String>() { "SyncAddress", configSyncAddress });
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
			} catch (Exception exception) {
				// Send message on error.
				this.SendMail("Error " + this.GetName(), exception.Message, false);

				// Throw the error.
				throw;
			}
		} // Run
		#endregion

	} // ActiveDirectorySynchronizeWithSofd
	#endregion

} // NDK.PluginCollection