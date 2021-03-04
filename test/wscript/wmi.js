  var wbemFlagReturnImmediately = 0x10;
var wbemFlagForwardOnly = 0x20;

var arrComputers = new Array(".");
for (i = 0; i < arrComputers.length; i++) {
   WScript.Echo();
   WScript.Echo("==========================================");
   WScript.Echo("Computer: " + arrComputers[i]);
   WScript.Echo("==========================================");

   var objWMIService = GetObject("winmgmts:\\\\" + arrComputers[i] + "\\root\\CIMV2");
   var colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem", "WQL",
                                          wbemFlagReturnImmediately | wbemFlagForwardOnly);

   var enumItems = new Enumerator(colItems);
   for (; !enumItems.atEnd(); enumItems.moveNext()) {
      var objItem = enumItems.item();

      WScript.Echo("AdminPasswordStatus: " + objItem.AdminPasswordStatus);
      WScript.Echo("AutomaticResetBootOption: " + objItem.AutomaticResetBootOption);
      WScript.Echo("AutomaticResetCapability: " + objItem.AutomaticResetCapability);
      WScript.Echo("BootOptionOnLimit: " + objItem.BootOptionOnLimit);
      WScript.Echo("BootOptionOnWatchDog: " + objItem.BootOptionOnWatchDog);
      WScript.Echo("BootROMSupported: " + objItem.BootROMSupported);
      WScript.Echo("BootupState: " + objItem.BootupState);
      WScript.Echo("Caption: " + objItem.Caption);
      WScript.Echo("ChassisBootupState: " + objItem.ChassisBootupState);
      WScript.Echo("CreationClassName: " + objItem.CreationClassName);
      WScript.Echo("CurrentTimeZone: " + objItem.CurrentTimeZone);
      WScript.Echo("DaylightInEffect: " + objItem.DaylightInEffect);
      WScript.Echo("Description: " + objItem.Description);
      WScript.Echo("DNSHostName: " + objItem.DNSHostName);
      WScript.Echo("Domain: " + objItem.Domain);
      WScript.Echo("DomainRole: " + objItem.DomainRole);
      WScript.Echo("EnableDaylightSavingsTime: " + objItem.EnableDaylightSavingsTime);
      WScript.Echo("FrontPanelResetStatus: " + objItem.FrontPanelResetStatus);
      WScript.Echo("InfraredSupported: " + objItem.InfraredSupported);
      try { WScript.Echo("InitialLoadInfo: " + (objItem.InitialLoadInfo.toArray()).join(",")); }
         catch(e) { WScript.Echo("InitialLoadInfo: null"); }
      WScript.Echo("InstallDate: " + WMIDateStringToDate(""+objItem.InstallDate));
      WScript.Echo("KeyboardPasswordStatus: " + objItem.KeyboardPasswordStatus);
      WScript.Echo("LastLoadInfo: " + objItem.LastLoadInfo);
      WScript.Echo("Manufacturer: " + objItem.Manufacturer);
      WScript.Echo("Model: " + objItem.Model);
      WScript.Echo("Name: " + objItem.Name);
      WScript.Echo("NameFormat: " + objItem.NameFormat);
      WScript.Echo("NetworkServerModeEnabled: " + objItem.NetworkServerModeEnabled);
      WScript.Echo("NumberOfProcessors: " + objItem.NumberOfProcessors);
      try { WScript.Echo("OEMLogoBitmap: " + (objItem.OEMLogoBitmap.toArray()).join(",")); }
         catch(e) { WScript.Echo("OEMLogoBitmap: null"); }
      try { WScript.Echo("OEMStringArray: " + (objItem.OEMStringArray.toArray()).join(",")); }
         catch(e) { WScript.Echo("OEMStringArray: null"); }
      WScript.Echo("PartOfDomain: " + objItem.PartOfDomain);
      WScript.Echo("PauseAfterReset: " + objItem.PauseAfterReset);
      try { WScript.Echo("PowerManagementCapabilities: " + (objItem.PowerManagementCapabilities.toArray()).join(",")); }
         catch(e) { WScript.Echo("PowerManagementCapabilities: null"); }
      WScript.Echo("PowerManagementSupported: " + objItem.PowerManagementSupported);
      WScript.Echo("PowerOnPasswordStatus: " + objItem.PowerOnPasswordStatus);
      WScript.Echo("PowerState: " + objItem.PowerState);
      WScript.Echo("PowerSupplyState: " + objItem.PowerSupplyState);
      WScript.Echo("PrimaryOwnerContact: " + objItem.PrimaryOwnerContact);
      WScript.Echo("PrimaryOwnerName: " + objItem.PrimaryOwnerName);
      WScript.Echo("ResetCapability: " + objItem.ResetCapability);
      WScript.Echo("ResetCount: " + objItem.ResetCount);
      WScript.Echo("ResetLimit: " + objItem.ResetLimit);
      try { WScript.Echo("Roles: " + (objItem.Roles.toArray()).join(",")); }
         catch(e) { WScript.Echo("Roles: null"); }
      WScript.Echo("Status: " + objItem.Status);
      try { WScript.Echo("SupportContactDescription: " + (objItem.SupportContactDescription.toArray()).join(",")); }
         catch(e) { WScript.Echo("SupportContactDescription: null"); }
      WScript.Echo("SystemStartupDelay: " + objItem.SystemStartupDelay);
      try { WScript.Echo("SystemStartupOptions: " + (objItem.SystemStartupOptions.toArray()).join(",")); }
         catch(e) { WScript.Echo("SystemStartupOptions: null"); }
      WScript.Echo("SystemStartupSetting: " + objItem.SystemStartupSetting);
      WScript.Echo("SystemType: " + objItem.SystemType);
      WScript.Echo("ThermalState: " + objItem.ThermalState);
      WScript.Echo("TotalPhysicalMemory: " + objItem.TotalPhysicalMemory);
      WScript.Echo("UserName: " + objItem.UserName);
      WScript.Echo("WakeUpType: " + objItem.WakeUpType);
      WScript.Echo("Workgroup: " + objItem.Workgroup);
   }
}

function WMIDateStringToDate(dtmDate)
{
   if (dtmDate == null || dtmDate=="null")
   {
      return "null date";
   }
   var strDateTime;
   if (dtmDate.substr(4, 1) == 0)
   {
      strDateTime = dtmDate.substr(5, 1) + "/";
   }
   else
   {
      strDateTime = dtmDate.substr(4, 2) + "/";
   }
   if (dtmDate.substr(6, 1) == 0)
   {
      strDateTime = strDateTime + dtmDate.substr(7, 1) + "/";
   }
   else
   {
      strDateTime = strDateTime + dtmDate.substr(6, 2) + "/";
   }
   strDateTime = strDateTime + dtmDate.substr(0, 4) + " " +
   dtmDate.substr(8, 2) + ":" +
   dtmDate.substr(10, 2) + ":" +
   dtmDate.substr(12, 2);
   return(strDateTime);
}