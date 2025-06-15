using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace SCS_SA_SDK
{
    [Serializable]
    public class SAInstrumentInterfaceParameters
    {
        private int serializeVersion = 2; // 2 = added powerLock boolean SCS 2024.03.06
        public SA_InstID instIdx;
        public bool initializeAtStartUp = true;// 
        public string deviceIPAddress = "";// Device IP Address (optional)
        public int interfaceType = 0; //Interface Type (0=default)
        public bool runInSimulation = false;
        public bool allowStartWithoutInitRequirements = false;
        public bool powerLock = true;

        public SAInstrumentInterfaceParameters()
        {
            instIdx = new SA_InstID();
        }

        public SAInstrumentInterfaceParameters(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                initializeAtStartUp = (bool)info.GetValue("initializeAtStartUp", typeof(bool)); 
                deviceIPAddress = info.GetString("deviceIPAddress");
                interfaceType = (int)info.GetValue("interfaceType", typeof(int));
                runInSimulation = (bool)info.GetValue("runInSimulation", typeof(bool));
                allowStartWithoutInitRequirements = (bool)info.GetValue("allowStartWithoutInitRequirements", typeof(bool));

                SA_InstID instIdx = new SA_InstID(info, ctxt);
            }
            if (sVersion >= 2)
            {
                powerLock = (bool)info.GetValue("powerLock", typeof(bool));
            }
        }

        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("initializeAtStartUp", initializeAtStartUp); 
            info.AddValue("deviceIPAddress",deviceIPAddress); // Device IP Address (optional)
            info.AddValue("interfaceType", interfaceType); //Interface Type (0=default)
            info.AddValue("runInSimulation", runInSimulation);
            info.AddValue("allowStartWithoutInitRequirements", allowStartWithoutInitRequirements);

            instIdx.GetObjectData(info, ctxt);

            // serializationVerion 2
            info.AddValue("powerLock", powerLock);
        }     
    }
}
