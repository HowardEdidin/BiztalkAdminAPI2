using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using BiztalkAdminAPI.Models;
using Microsoft.AspNetCore.Mvc;

namespace BiztalkAdminAPI.Controllers
{
    [Produces("application/json")]
    [Route("[controller]")]
    [ApiController]
    public class HostInstanceController : ControllerBase
    {
        /// <summary>
        /// Lists host instances in BizTalk server
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IEnumerable<HostInstance> Get()
        {
            List<HostInstance> hostInstances = new List<HostInstance>();
            try
            {
                //Create EnumerationOptions and run wql query 
                EnumerationOptions enumOptions = new EnumerationOptions
                {
                    ReturnImmediately = false
                };

                //Search for all HostInstances 
                ManagementObjectSearcher searchObject = new ManagementObjectSearcher("root\\MicrosoftBizTalkServer", "Select * from MSBTS_HostInstance", enumOptions);

                //Enumerate through the resultset 
                foreach (ManagementObject inst in searchObject.Get())
                {
                    HostInstance hostInstance = new HostInstance
                    {
                        Caption = inst["Caption"] != null ? inst["Caption"].ToString() : "",
                        ClusterInstanceType = inst["ClusterInstanceType"] != null ? inst["ClusterInstanceType"].ToString() : "",
                        ConfigurationState = inst["ConfigurationState"] != null ? inst["ConfigurationState"].ToString() : "",
                        Description = inst["Description"] != null ? inst["Description"].ToString() : "",
                        HostName = inst["HostName"] != null ? inst["HostName"].ToString() : "",
                        HostType = inst["HostType"] != null ? Enum.GetName(typeof(HostInstance.HostTypeEnum), inst["HostType"]) : "",
                        InstallDate = inst["InstallDate"] != null ? inst["InstallDate"].ToString() : "",
                        IsDisabled = inst["IsDisabled"] != null ? inst["IsDisabled"].ToString() : "",
                        Logon = inst["Logon"] != null ? inst["Logon"].ToString() : "",
                        MgmtDbNameOverride = inst["MgmtDbNameOverride"] != null ? inst["MgmtDbNameOverride"].ToString() : "",
                        MgmtDbServerOverride = inst["MgmtDbServerOverride"] != null ? inst["MgmtDbServerOverride"].ToString() : "",
                        Name = inst["Name"] != null ? inst["Name"].ToString() : "",
                        NTGroupName = inst["NTGroupName"] != null ? inst["NTGroupName"].ToString() : "",
                        RunningServer = inst["RunningServer"] != null ? inst["RunningServer"].ToString() : "",
                        ServiceState = inst["ServiceState"] != null ? Enum.GetName(typeof(HostInstance.ServiceStateEnum), inst["ServiceState"]) : "",
                        Status = inst["Status"] != null ? inst["Status"].ToString() : "",
                        UniqueID = inst["UniqueID"] != null ? inst["UniqueID"].ToString() : ""
                    };

                    hostInstances.Add(hostInstance);
                }
            }
            catch (Exception ex)
            {
                //System.Diagnostics.EventLog.WriteEntry("BizTalk Server", "Exception Occurred in enumerateAndStartHostInstances fuction call. " + excep.Message, System.Diagnostics.EventLogEntryType.Error);
                throw new Exception("Exception Occurred in get hostinstances call. " + ex.Message);
            }
            return hostInstances;
        }

        /// <summary>
        /// Start or stop host instance.
        /// </summary>
        /// <remarks>
        /// Sample requests:
        ///  PUT /BizTalkAdminAPI/HostInstance/servername/hostinstance/Start
        ///  PUT /BizTalkAdminAPI/HostInstance/servername/hostinstance/Stop
        /// </remarks>
        /// <param name="servername"></param>
        /// <param name="hostname"></param>
        /// <param name="state"></param>
        [HttpPut("{servername}/{hostname}/{state}")]
        public void Put(string servername, string hostname, string state)
        {
            if (state == "Start" || state == "Stop")
            {
                EnumerationOptions enumOptions = new EnumerationOptions
                {
                    ReturnImmediately = false
                };

                ManagementObjectSearcher searchObject = new ManagementObjectSearcher("root\\MicrosoftBizTalkServer", "Select * from MSBTS_HostInstance Where RunningServer='" + servername + "' And HostName='" + hostname + "'", enumOptions);

                if (searchObject.Get().Count > 0)
                {
                    ManagementObjectCollection instcol = searchObject.Get();
                    ManagementObject inst = instcol.OfType<ManagementObject>().FirstOrDefault();

                    inst.InvokeMethod(state, null);
                }
            }
        }

        
    }
}
