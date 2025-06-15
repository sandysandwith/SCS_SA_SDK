using SpatialAnalyzerSDK;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using static System.Net.WebRequestMethods;

namespace SCS_SA_SDK
{
    public class SA_sdk : SpatialAnalyzerSDKClass
    {
        public bool connected = false;
        int statusCode = 0;
        public bool busySA = false;

        public enum MPStatus
        {
            SdkError = -1,
            Undone = 0,
            InProgress = 1,
            DoneSuccess = 2,
            DoneFatalError = 3,
            DoneMinorError = 4,
            CurrentTask = 5
        };

        private bool IsDoneSuccess()
        {
            int rCode = 0;
            GetMPStepResult(ref rCode);
            if (rCode != (int)MPStatus.DoneSuccess)
                return false;
            return true;
        }
        
        public bool ConnectToSA(string host = "localhost")
        {
            connected = ConnectEx(host, ref statusCode);  // what values of status code are possible? 
            if (!connected || statusCode != 0)
            {
                Console.WriteLine("Unable to Connect to SA.  Status Code: " + statusCode);
            }
            else
            {
                SetInteractionMode();
            }
            return connected;
        }

        public bool DissconnectFromSA()
        {
            bool ret = false;
            if (connected)
            {
                //ReleaseDispatch();

            }

            return ret;
        }

        public bool ShutDownSA()
        {
            SetStep("Shut Down SA");
            bool ret = ExecuteStep();

            return ret;
        }

        public bool SetInteractionMode()
        {
            SetStep("Set Interaction Mode");
            // Available options:  
            // "Manual", "Automatic", "Silent", C:\Users\Scott\Documents\Visual Studio 2010\Projects\SA_CADVal_App\SA_DPD_TestPlatform\mySDK.cs
            SetSAInteractionModeArg("SA Interaction Mode", "Silent");
            // Available options: 
            // "Halt on Failure Only", "Halt on Failure or Partial Success", "Never Halt", 
            SetMPInteractionModeArg("Measurement Plan Interaction Mode", "Never Halt");
            // Available options: 
            // "Block Application Interaction", "Allow Application Interaction", 
            SetMPDialogInteractionModeArg("Measurement Plan Dialog Interaction Mode", "Allow Application Interaction");
            ExecuteStep();

            return IsDoneSuccess();
        }




        public bool NewSAFile()
        {
            SetStep("New SA File");
            bool ret = ExecuteStep();

            return IsDoneSuccess();
        }

        public bool SAFileSaveAs(string fn)
        {
            SetStep("Save As...");
            SetFilePathArg("File Name", fn, false);
            SetBoolArg("Add Serial Number?", false);
            SetIntegerArg("Optional Number", 0);
            ExecuteStep();
            
            return IsDoneSuccess();
        }

        public bool OpenSAFile(string filename)
        {
            SetStep("Open SA File");
            SetFilePathArg("SA File Name", filename, false);
            ExecuteStep();
            
            return IsDoneSuccess();
        }

        public bool OpenSATemplateFile(string filename)
        {
            SetStep("Open Template File");
            SetFilePathArg("Template File Name", filename, false);
            ExecuteStep();
 
            return IsDoneSuccess();
        }


        public bool RunSubroutine(string sub_FileName)
        {
            SetStep("Run Subroutine");
            SetFilePathArg("MP Subroutine File Path", sub_FileName, false);
            SetBoolArg("Share Parent Variables?", true);
            ExecuteStep();
            
            return IsDoneSuccess();
        }

        public bool MakeACollectionObjectNameRefList_ByType(String coll, SA_Type type, ref List<String> strList)
        {
            SetStep("Make a Collection Object Name Ref List - By Type");
            SetStringArg("Collection", coll);
            // Available options: 
            // "Any", "B-Spline", "Circle", "Cloud", "Scan Stripe Cloud", 
            // "Cross Section Cloud", "Cone", "Cylinder", "Datum", "Ellipse", 
            // "Frame", "Frame Set", "Line", "Paraboloid", "Perimeter", 
            // "Plane", "Point Group", "Poly Surface", "Scan Stripe Mesh", "Slot", 
            // "Sphere", "Surface", "Vector Group", 
            SetObjectTypeArg("Object Type", type.GetTypeName());
            bool ret = ExecuteStep();

            if (ret)
            {
                if (IsDoneSuccess())
                {
                    GetCollectionObjectRefListFromSA("Resultant Collection Object Name List", ref strList);
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        public bool StartInstrumentInterface(SAInstrumentInterfaceParameters iip)
        {
            bool ret = false;
            bool connected = false;
            ret = VerifyInstrumentConnection(iip, ref connected);
            if (!ret)
            {
                SetStep("Start Instrument Interface");
                SetColInstIdArg("Instrument's ID", iip.instIdx.GetColl(), iip.instIdx.GetInstIdx());
                SetBoolArg("Initialize at Startup", true);
                SetStringArg("Device IP Address (optional)", iip.deviceIPAddress);
                SetIntegerArg("Interface Type (0=default)", iip.interfaceType);
                SetBoolArg("Run in Simulation", iip.runInSimulation);
                SetBoolArg("Allow Start w/o Init Requirements", iip.allowStartWithoutInitRequirements);
                ExecuteStep();
                ret = VerifyInstrumentConnection(iip, ref connected);
            }

            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode != (int)MPStatus.DoneSuccess)
                {
                    connected = false;
                }
            }
            return connected;

        }

        public bool StopInstrumentInterface(SAInstrumentInterfaceParameters iip)
        {
            bool connected = true;
            VerifyInstrumentConnection(iip, ref connected);
            if (!connected)
                return true;

            /// instrument is connected so disconnect
            /// 
            SetStep("Stop Instrument Interface");
            SetColInstIdArg("Instrument's ID", iip.instIdx.GetColl(), iip.instIdx.GetInstIdx());
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode != (int)MPStatus.DoneSuccess)
                {
                    ret = false;
                }
            }
            return ret;
        }


        public bool ResetInstrumentInterface(SAInstrumentInterfaceParameters iip)
        {
            bool InstConnected = false;
            VerifyInstrumentConnection(iip, ref InstConnected);
            if (InstConnected)
            {
                StopInstrumentInterface(iip);
                Thread.Sleep(1000);
                StartInstrumentInterface(iip);
                Thread.Sleep(1000);
            }
            VerifyInstrumentConnection(iip, ref InstConnected);
            return InstConnected;
        }
        public bool GetInstrumentWeatherSettings(SA_InstID instId, ref double temperature, ref double pressure, ref double humidity, ref bool setAutomatically)
        {
            SetStep("Get Instrument Weather Setting");
            SetColInstIdArg("Instrument's ID", instId.GetColl(), instId.instIdx);
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetDoubleArg("Temperature (F)", ref temperature);
                    GetDoubleArg("Pressure (mmHg)", ref pressure);
                    GetDoubleArg("Humidity (%Rel)", ref humidity);
                    GetBoolArg("Was Set Automatically? (using Inst or external sensor", ref setAutomatically);
                }
                else
                {
                    ret = false;
                }
            }

            return ret;
        }

        public bool SetInstrumentWeatherSettings(SA_InstID instId, double temperature, double pressure, double humidity)
        {
            SetStep("Set Instrument Weather Setting");
            SetColInstIdArg("Instrument's ID", instId.GetColl(), instId.instIdx);
            SetDoubleArg("Temperature (F)", temperature);
            SetDoubleArg("Pressure (mmHg)", pressure);
            SetDoubleArg("Humidity (%Rel)", humidity);
            SetBoolArg("Set Automatically? (Ignore above values)", false);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool GetActiveUnits(ref string sLenghtUnits, ref string sAngularUnits, ref string sTemperatureUnits)
        {
            SetStep("Get Active Units");
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    ret = GetStringArg("Length", ref sLenghtUnits);
                    if (ret)
                        ret = GetStringArg("Angular", ref sAngularUnits);
                    if (ret)
                        ret = GetStringArg("Temperature", ref sTemperatureUnits);
                }
                else
                    ret = false;
            }
            return ret;
        }

        public bool SetActiveUnits(string sLengthUnits, bool bDisplayInchFraction, double dFractionDenominator, bool bSimplifyFractions, string sTemperatureUnits, string sAngularUnits)
        {
            SetStep("Set Active Units");
            // Available options: 
            // "Meters", "Centimeters", "Millimeters", "Feet", "Inches", 
            SetDistanceUnitsArg("Length", sLengthUnits);
            SetBoolArg("Display Inch Fractions?", bDisplayInchFraction);
            SetDoubleArg("Inch Fraction Denominator?", dFractionDenominator);
            SetBoolArg("Simplify Inch Fraction?", bSimplifyFractions);
            // Available options: 
            // "Fahrenheit", "Celsius", 
            SetTemperatureUnitsArg("Temperature", sTemperatureUnits);
            // Available options: 
            // "Degrees", "Deg:Min:Sec", "Radians", "Milliradians", 
            // "Gons/Grad", "Mils", "Arcseconds", "Deg:Min", 
            SetAngularUnitsArg("Angular", sAngularUnits);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool GetMeasurementInfoData(SA_TargetName targ, ref string infoData)
        {
            SetStep("Get Measurement Info Data");
            SetPointNameArg("Point Name", targ.GetCollName(), targ.GetGrpName(), targ.GetTargetName());
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetStringArg("Info Data", ref infoData);
                }
                else
                    ret = false;
            }
            return ret;
        }


        public bool VerifyInstrumentConnection(SAInstrumentInterfaceParameters iip, ref bool instConnected)
        {
            SetStep("Verify Instrument Connection");
            SetColInstIdArg("Instrument's ID", iip.instIdx.GetColl(), iip.instIdx.GetInstIdx());
            bool ret = ExecuteStep();

            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetBoolArg("Connected?", ref instConnected);
                    if (instConnected)
                        ret = true;
                }
                else
                    ret = false;
            }
            return ret;
        }

        public bool GetLastInstrumentIndex(ref SA_InstID inst)
        {
            SetStep("Get Last Instrument Index");
            bool ret = ExecuteStep();

            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetIntegerArg("Instrument ID", ref inst.instIdx);
                    GetColInstIdArg("Instrument ID", ref inst.Coll, ref inst.instIdx);
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        public bool GetInstrumentTransform(SA_InstID inst, SA_CollObjectName Frame, ref Transform t)
        {
            SetStep("Get Instrument Transform");
            SetColInstIdArg("Instrument ID", inst.GetColl(), inst.instIdx);
            SetCollectionObjectNameArg("Reference Frame", Frame.GetCollName(), Frame.GetObjName());
            bool ret = ExecuteStep();
            if (ret)
            {
                ret = GetTransformFromSA("Transform", ref t);
                if (ret)
                {
                    int rCode = 0;
                    GetMPStepResult(ref rCode);
                    if (rCode != (int)MPStatus.DoneSuccess)
                    {
                        ret = false;
                    }
                }
            }
            return ret;
        }

        public bool SetInstrumentTransformToLastInsturment(SA_InstID oldInst, SA_CollObjectName Frame, ref SA_InstID newInst)
        {
            Transform t = new Transform();
            SA_CollObjectName workingFrame = new SA_CollObjectName();
            bool ret = GetInstrumentTransform(oldInst, Frame, ref t);
            if (ret)
            {
                ret = GetWorkingFrameProperties(ref workingFrame);
                if (ret)
                {
                    ret = SetInstrumentTransform(newInst, t, workingFrame);
                }
            }
            return ret;
        }

        public bool SetInstrumentTransform(SA_InstID inst, Transform t, SA_CollObjectName referenceFrame)
        {
            SetStep("Set Instrument Transform");
            SetColInstIdArg("Instrument to Move", inst.GetColl(), inst.instIdx);
            SetInputTransformArg("Destination Transform", t);

            SetCollectionObjectNameArg("Reference Frame", referenceFrame.GetCollName(), referenceFrame.GetObjName());
            SetIntegerArg("Number of Steps", 0);
            ExecuteStep();

            return IsDoneSuccess();
        }

        public bool GetNumberOfCollections(ref int nCollections)
        {
            SetStep("Get Number of Collections");
            bool ret = ExecuteStep();

            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetIntegerArg("Total Count", ref nCollections);
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        public bool SetActiveCollection(string collName)
        {
            SetStep("Set (or construct) default collection");
            SetCollectionNameArg("Collection Name", collName);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool GetithCollectionName(int i, ref string collName)
        {
            SetStep("Get i-th Collection Name");
            SetIntegerArg("Collection Index", i);
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetCollectionNameArg("Resultant Name", ref collName);
                }
                else
                { 
                    ret = false; 
                }
            }
            return ret;
        }

        public bool GetListOfCollections(ref List<string> listOfCollections)
        {
            int nCollections = 0;

            bool ret = GetNumberOfCollections(ref nCollections);
            if (ret)
            {
                for (int i = 0; i < nCollections; i++)
                {
                    string collName = "";
                    if (GetithCollectionName(i, ref collName))
                    {
                        listOfCollections.Add(collName);
                    }
                }
            }
            return ret;
        }

        public bool GetActiveCollectionName(ref string collName)
        {
            SetStep("Get Active Collection Name");
            bool ret = ExecuteStep();

            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetCollectionNameArg("Currently Active Collection Name", ref collName);
                }
                else
                {
                    ret= false;
                }
            }
            return ret;
        }

        public bool GetInstrumentModel(ref SA_InstID inst)
        {
            SetStep("Get Instrument Model");
            SetColInstIdArg("Instrument ID", inst.GetColl(), inst.instIdx);
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetStringArg("Name", ref inst.Name);
                    GetStringArg("Model", ref inst.Model);
                }
                else
                { 
                    ret= false; 
                }
            }
            return ret;
        }

        public bool AutoMeasurePoints(SAInstrumentInterfaceParameters iip, SA_CollObjectName nom_ptgrp, SA_CollObjectName meas_ptgrp)
        {
            bool connected = true;
            VerifyInstrumentConnection(iip, ref connected);
            if (!connected)
            {
                connected = StartInstrumentInterface(iip);
            }
            if (connected)
            {
                SetStep("Auto Measure Points");
                SetColInstIdArg("Instrument ID", iip.instIdx.GetColl(), iip.instIdx.GetInstIdx());
                SetCollectionObjectNameArg("Reference Group Name", nom_ptgrp.GetCollName(), nom_ptgrp.GetObjName());
                SetCollectionObjectNameArg("Actuals Group Name (to be measured)", meas_ptgrp.GetCollName(), meas_ptgrp.GetObjName());
                SetBoolArg("Force use of existing group?", false);
                SetBoolArg("Show complete dialog?", false);
                SetBoolArg("Wait for Completion?", true);
                SetBoolArg("Auto Start?", true);
                ExecuteStep();
            }

            return IsDoneSuccess();
        }


        public bool GetCurrentInstrumentPositionUpdate(SA_InstID inst, string ReportingFrame, ref Vector3D xyz, ref double time_since_update_sec, ref string timeStamp)
        {
            SetStep("Get Current Instrument Position Update");
            SetColInstIdArg("Instrument ID", inst.GetColl(), inst.instIdx);
            // Available options: 
            // "Instrument Base", "World", "Working", 
            SetStringArg("Reporting Frame", ReportingFrame);
            SetBoolArg("Polar Coordinates?", false);
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetDoubleArg(@"X / R", ref xyz.x);
                    GetDoubleArg(@"Y / Theta (Degrees)", ref xyz.y);
                    GetDoubleArg(@"Z / Phi (Degrees)", ref xyz.z);
                    GetDoubleArg(@"Time Since Update (sec)", ref time_since_update_sec);
                    GetStringArg(@"Timestamp (Approximate)", ref timeStamp);
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        public bool LocateInstrumentBestFitGroupToGroup(ToleranceSet ts, SA_CollObjectName nom_ptgrp, SA_CollObjectName meas_ptgrp, ref Transform T, bool genEvents = false)
        {
            SetStep("Locate Instrument (Best Fit - Group to Group)");
            SetCollectionObjectNameArg("Reference Group", nom_ptgrp.GetCollName(), nom_ptgrp.GetObjName());
            SetCollectionObjectNameArg("Corresponding Group", meas_ptgrp.GetCollName(), meas_ptgrp.GetObjName());
            SetBoolArg("Show Interface", false);
            SetDoubleArg("RMS Tolerance (0.0 for none)", ts.rmsTol);
            SetDoubleArg("Maximum Absolute Tolerance (0.0 for none)", ts.maxTol);
            SetBoolArg("Allow Scale", false);
            SetBoolArg("Allow X", true);
            SetBoolArg("Allow Y", true);
            SetBoolArg("Allow Z", true);
            SetBoolArg("Allow Rx", true);
            SetBoolArg("Allow Ry", true);
            SetBoolArg("Allow Rz", true);
            SetBoolArg("Lock Degrees of Freedom", false);
            SetBoolArg("Generate Event", genEvents);
            SetFilePathArg("File Path for CSV Text Report (requires Show Interface = true)", "", false);
            bool ret = ExecuteStep();
            if (ret)
            {
                GetTransformInWorkingFromSA(ref T);

                Transform WT = new Transform();

                double scale = 1.0;
                GetWorldTransformOptimumTransformFromSA(ref WT, ref scale);
                GetDoubleArg("RMS Deviation", ref ts.rmsResult);
                GetDoubleArg("Maximum Absolute Deviation", ref ts.maxResult);

            }

            return IsDoneSuccess();
        }

        public bool SetWorkingFrame(SA_CollObjectName workingFrame)
        {
            SetStep("Set Working Frame");
            SetCollectionObjectNameArg("New Working Frame Name", workingFrame.GetCollName(), workingFrame.GetObjName());
            ExecuteStep();

            return IsDoneSuccess();
        }

        public bool GetWorkingFrameProperties(ref SA_CollObjectName workingFrame)
        {
            SetStep("Get Working Frame Properties");
            bool ret = ExecuteStep();

            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    ret = GetStringArg("Frame Name", ref workingFrame.objectName);
                    if (ret)
                        ret = GetStringArg("Collection Name", ref workingFrame.Coll);
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        public bool JumpInstrumentToNewLocation(SA_InstID instIdx)
        {
            SetStep("Jump Instrument To New Location");
            SetColInstIdArg("Live Instrument ID", instIdx.GetColl(), instIdx.instIdx);
            SetBoolArg("Hide the Previous Instrument?", true);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool BestFitTransformationGroupToGroup(ref ToleranceSet ts, SA_CollObjectName nom_ptgrp, SA_CollObjectName meas_ptgrp, DoF dof, ref Transform T, bool genEvent = false)
        {
            SetStep("Best Fit Transformation - Group to Group");
            SetCollectionObjectNameArg("Reference Group", nom_ptgrp.GetCollName(), nom_ptgrp.GetObjName());
            SetCollectionObjectNameArg("Corresponding Group", meas_ptgrp.GetCollName(), meas_ptgrp.GetObjName());
            SetBoolArg("Show Interface", false);
            SetDoubleArg("RMS Tolerance (0.0 for none)", ts.rmsTol);
            SetDoubleArg("Maximum Absolute Tolerance (0.0 for none)", ts.maxTol);
            SetBoolArg("Allow Scale", dof.dofScale);
            SetBoolArg("Allow X", dof.dofX);
            SetBoolArg("Allow Y", dof.dofY);
            SetBoolArg("Allow Z", dof.dofZ);
            SetBoolArg("Allow Rx", dof.dofRx);
            SetBoolArg("Allow Ry", dof.dofRy);
            SetBoolArg("Allow Rz", dof.dofRz);
            SetBoolArg("Lock Degrees of Freedom", false);
            SetBoolArg("Generate Event", genEvent);
            SetFilePathArg("File Path for CSV Text Report (requires Show Interface = true)", "", false);
            bool ret = ExecuteStep();

            if (ret)
            {
                GetTransformInWorkingFromSA(ref T);

                Transform WT = new Transform();

                double scale = 1.0;
                GetWorldTransformOptimumTransformFromSA(ref WT, ref scale);
                GetDoubleArg("RMS Deviation", ref ts.rmsResult);
                GetDoubleArg("Maximum Absolute Deviation", ref ts.maxResult);
            }
            return IsDoneSuccess();
        }

        public bool SetLeicaTrackerPowerLock(SA_InstID instIdx, bool lockState)
        {
            string sLockState = "AutoLock On";
            if (lockState == false)
                sLockState = "AutoLock Off";

            bool ret = InstrumentOperationalCheck(instIdx, sLockState);

            return ret;
        }

        public bool MakeACollectionInstrumentRefList(String coll, ref List<String> instIdList)
        {
            bool ret = SetActiveCollection(coll);
            if (ret)
            {
                SetStep("Make a Collection Instrument Reference List");
                ret = ExecuteStep();

                if (ret)
                {
                    GetColInstRefListFromSA(ref instIdList);
                }
            }
            return IsDoneSuccess();
        }

        private bool GetColInstRefListFromSA(ref List<String> instIdList)
        {
            Object o = null;
            VariantWrapper vlist = new VariantWrapper(o);
            Object il = vlist;

            bool ret = GetColInstIdRefListArg("Resultant Collection Instrument Reference List", ref il);

            if (ret)
            {
                if (IsDoneSuccess())
                {
                    Array a_il = (Array)il;

                    for (int i = a_il.GetLowerBound(0); i <= a_il.GetUpperBound(0); i++)
                    {
                        String s = (String)a_il.GetValue(i);
                        instIdList.Add(s);
                    }
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        private bool GetColInstIdxFromSA(ref SA_InstID inst)
        {
            bool ret = GetColInstIdArg("Instrument Added (result)", ref inst.Coll, ref inst.instIdx);
            return ret;
        }

        private bool GetCollectionObjectRefListFromSA(string argName, ref List<String> strList)
        {
            Object o = null;
            VariantWrapper str_list = new VariantWrapper(o);
            Object sl = str_list;

            bool ret = GetCollectionGroupNameRefListArg(argName, ref sl);

            if (ret)
            {
                if (IsDoneSuccess())
                {
                    Array a_sl = (Array)sl;

                    for (int i = a_sl.GetLowerBound(0); i <= a_sl.GetUpperBound(0); i++)
                    {
                        String s = (String)a_sl.GetValue(i);
                        strList.Add(s);
                    }
                }
                else
                { 
                ret = false;
                }
            }
            return ret;
        }



        public bool MakeAPointNameRefListFromAGroup(SA_CollObjectName ptGrp, ref SA_PointNameRefList ptNameList)
        {
            SetStep("Make a Point Name Ref List From a Group");
            SetCollectionObjectNameArg("Group Name", ptGrp.GetCollName(), ptGrp.GetObjName());
            bool ret = ExecuteStep();

            if (ret)
            {
                ret = GetPointRefListFromSA("Resultant Point Name List", ref ptNameList.ptNameRefList);
            }
            else
            {
                ret = false; 
            }
            return ret;
        }

        private bool GetPointRefListFromSA(string argName, ref List<SA_TargetName> ptNameList)
        {
            Object o = null;
            VariantWrapper str_list = new VariantWrapper(o);
            Object sl = str_list;

            bool ret = GetPointNameRefListArg(argName, ref sl);

            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    Array a_sl = (Array)sl;

                    for (int i = a_sl.GetLowerBound(0); i <= a_sl.GetUpperBound(0); i++)
                    {
                        String s = (String)a_sl.GetValue(i);
                        SA_TargetName tn = new SA_TargetName(s);
                        ptNameList.Add(tn);
                    }
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        private bool GetStringRefListFromSA(string argName, ref List<string> sList)
        {
            Object o = null;
            VariantWrapper str_list = new VariantWrapper(o);
            Object sl = str_list;

            bool ret = GetStringRefListArg(argName, ref sl);

            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    Array a_sl = (Array)sl;

                    for (int i = a_sl.GetLowerBound(0); i <= a_sl.GetUpperBound(0); i++)
                    {
                        String s = (String)a_sl.GetValue(i);
                        sList.Add(s);
                    }
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        public bool InstrumentOperationalCheck(SA_InstID inst, string OpCheckTypeName)
        {
            SetStep("Instrument Operational Check");
            SetColInstIdArg("Instrument to Check", inst.GetColl(), inst.instIdx);
            SetStringArg("Check Type", OpCheckTypeName);
            bool ret = ExecuteStep();
            int rCode = 0;

            if (ret)
            {
                GetMPStepResult(ref rCode);
                if (rCode != (int)MPStatus.DoneSuccess)
                {
                    ret = false;
                }
            }
            return ret;
        }



        public bool TransformObjectsByDeltaWorldTransformOperator(SA_CollObjectNameList conl, Transform T)
        {
            Object vObjList = null;
            List<string> sl = new List<string>();
            conl.GetStringList(ref sl);
            MakeRefObjectList(sl, ref vObjList);
            SetCollectionObjectNameRefListArg("Objects to Transform", ref vObjList);

            SetStep("Transform Objects by Delta (World Transform Operator)");
            double scale = 1.0;
            SetWorldTransformArg(T, scale);
            ExecuteStep();

            return IsDoneSuccess();
        }

        public bool TransformObjectsByDeltaAboutWorkingTransformOperator(SA_CollObjectNameList conl, Transform T)
        {
            Object vObjList = null;
            List<string> sl = new List<string>();
            conl.GetStringList(ref sl);
            SetStep("Transform Objects by Delta (About Working Frame)");
            MakeRefObjectList(sl, ref vObjList);
            SetCollectionObjectNameRefListArg("Objects to Transform", ref vObjList);
            SetDeltaTransformArg(T);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool TransformInstrumentByDelta(SAInstrumentInterfaceParameters iip, Transform T)
        {
            SetStep("Transform Instrument by Delta");
            SetColInstIdArg("Instrument to Transform", iip.instIdx.GetColl(), iip.instIdx.GetInstIdx());
            SetDeltaTransformArg(T);
            SetBoolArg("Apply Scale from Transform to Instrument", false);
            ExecuteStep();
            return IsDoneSuccess();
        }

        private bool GetTransformInWorkingFromSA(ref Transform T)
        {
            VariantWrapper transform = new VariantWrapper(T.T);
            Object t = transform;

            bool ret = GetTransformArg("Transform in Working", ref t);
            CopyTransform(ref T, t);

            return ret;
        }

        private bool GetResultantTransformInWorkingFromSA(ref Transform T)
        {
            VariantWrapper transform = new VariantWrapper(T.T);
            Object t = transform;

            bool ret = GetTransformArg("Resultant Transform", ref t);
            CopyTransform(ref T, t);

            return ret;
        }


        private bool GetTransformFromSA(ref Transform T)
        {
            VariantWrapper transform = new VariantWrapper(T.T);
            Object t = transform;

            bool ret = GetTransformArg("Transform", ref t);
            CopyTransform(ref T, t);

            return ret;
        }

        private void CopyTransform(ref Transform T, Object t)
        {
            Array a_t = (Array)t;
            T.T = (double[,])a_t;
        }

        private bool GetWorldTransformOptimumTransformFromSA(ref Transform T, ref double s)
        {
            VariantWrapper transform = new VariantWrapper(T.T);
            Object t = transform;

            bool ret = GetWorldTransformArg("Optimum Transform", ref t, ref s);

            CopyTransform(ref T, t);

            return ret;
        }

        public bool GetWorkingTransformOfObject(SA_CollObjectName o, ref Transform T)
        {
            SetStep(@"Get Working Transform of Object (Fixed XYZ)");
            SetCollectionObjectNameArg("Object Name", o.GetCollName(), o.GetObjName());
            bool ret = ExecuteStep();
            if (ret)
            {
                ret = GetTransformFromSA(ref T);
            }
            return ret;
        }

        private bool SetInputTransformArg(Transform T)
        {
            bool ret = false;
            Object vTransform = new VariantWrapper(T.T);
            try
            {
                ret = SetTransformArg("Input Transform", ref vTransform);
            }
            catch (Exception ex)
            {
                // MessageBox.Show("Could not set the transform in SA: " + ex.ToString());
                Trace.WriteLine("Could not set the transform in SA: " + ex.ToString());
            }

            return ret;
        }

        public bool MakeATransformFromDoubles(Vector3D d, Vector3D r, ref Transform T)
        {
            SetStep("Make a Transform from Doubles (Fixed XYZ)");
            SetDoubleArg("X", d.x);
            SetDoubleArg("Y", d.y);
            SetDoubleArg("Z", d.z);
            SetDoubleArg("Rx (Roll)", r.x);
            SetDoubleArg("Ry (Pitch)", r.y);
            SetDoubleArg("Rz (Yaw)", r.z);
            bool ret = ExecuteStep();

            if(ret)
            {
                ret = GetResultantTransformInWorkingFromSA(ref T);
            }
            return ret;
        }

        private bool SetInputTransformArg(string argName, Transform T)
        {
            bool ret = false;
            Object vTransform = new VariantWrapper(T.T);
            try
            {
                ret = SetTransformArg(argName, ref vTransform);
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Could not set the transform in SA: " + ex.ToString());
                Trace.Write("Could not set the transform in SA: " + ex.ToString() + "/r/n");
            }

            return ret;
        }


        private bool SetDeltaTransformArg(Transform T)
        {
            bool ret = false;
            Object vTransform = new VariantWrapper(T.T);
            try
            {
                ret = SetTransformArg("Delta Transform", ref vTransform);
            }
            catch (Exception ex)
            {
                // MessageBox.Show("Could not set the transform in SA: " + ex.ToString());
                Trace.Write("Could not set the transform in SA: " + ex.ToString() + "/r/n");
            }

            return ret;
        }

        private bool SetWorldTransformArg(Transform T, double scale)
        {
            bool ret = false;

            Object vTransform = new VariantWrapper(T.T);
            try
            {
                ret = SetWorldTransformArg("Delta Transform", ref vTransform, scale);
            }
            catch (Exception ex)
            {
                // MessageBox.Show("Could not set the transform in SA: " + ex.ToString());
                Trace.Write("Could not set the transform in SA: " + ex.ToString() + "/r/n");
            }

            return ret;
        }

        public bool InvertTransform(Transform T, ref Transform InvertedT)
        {
            SetStep("Invert Transform");
            SetInputTransformArg("Transform", T);
            bool ret = ExecuteStep();

            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    ret = GetTransformFromSA("Inverse Transform", ref InvertedT);
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }


        private bool GetTransformFromSA(string arg, ref Transform T)
        {
            VariantWrapper transform = new VariantWrapper(T.T);
            Object t = transform;

            bool ret = GetTransformArg(arg, ref t);
            CopyTransform(ref T, t);

            return ret;
        }

        public bool DecomposeTransformIntoDoublesEulerXYZ(Transform T, ref Vector3D xyz, ref Vector3D Rxyz)
        {
            SetStep(@"Decompose Transform into Doubles (Euler XYZ)");
            bool ret = SetInputTransformArg(T);
            if (ret)
            {
                ret = ExecuteStep();

                if (ret)
                {
                    int rCode = 0;
                    GetMPStepResult(ref rCode);
                    if (rCode == (int)MPStatus.DoneSuccess)
                    {
                        GetDoubleArg("X", ref xyz.x);
                        GetDoubleArg("Y", ref xyz.y);
                        GetDoubleArg("Z", ref xyz.z);
                        GetDoubleArg("Euler Rx", ref Rxyz.x);
                        GetDoubleArg("Euler Ry", ref Rxyz.y);
                        GetDoubleArg("Euler Rz", ref Rxyz.z);
                    }
                    else
                    {
                        ret = false;
                    }
                }

            }
            return ret;
        }

        public bool ConstructFrame(SA_CollObjectName f, Transform T)
        {
            SetStep("Construct Frame");
            SetCollectionObjectNameArg("New Frame Name", f.GetCollName(), f.GetObjName());
            SetInputTransformArg("Transform in Working Coordinates", T);
            ExecuteStep();

            return IsDoneSuccess();

        }

        public bool CopyObject(SA_CollObjectName s, SA_CollObjectName d, bool overwrite)
        {
            SetStep("Copy Object");
            SetCollectionObjectNameArg("Source Object", s.GetCollName(), s.GetObjName());
            SetCollectionObjectNameArg("New Object Name", d.GetCollName(), d.GetObjName());
            SetBoolArg("Overwrite if exists?", overwrite);
            ExecuteStep();

            return IsDoneSuccess();
        }

        public bool ObjectExistanceTestCheckOnly(SA_CollObjectName o)
        {
            SetStep("Object Existence Test(Check Only)");
            SetCollectionObjectNameArg("Object Name", o.GetCollName(), o.GetObjName());

            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    bool exists = false;
                    GetBoolArg("Exists?", ref exists);
                    return exists;
                }
                else
                {
                    ret = false;
                }
            }
            return false;
        }



        public bool SetActiveUnits(SA_Units u)
        {
            SetStep("Set Active Units");
            // Available options: 
            // "Meters", "Centimeters", "Millimeters", "Feet", "Inches", 
            SetDistanceUnitsArg("Length", u.GetLengthUnit());
            // NOT_SUPPORTED("Temperature");
            // Available options: 
            // "Degrees", "Deg:Min:Sec", "Radians", "Milliradians", "Gons/Grad", "Mils", "Arcseconds", "Deg:Min", 
            SetAngularUnitsArg("Angular", u.GetAngleUnit());
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool DeleteObjects(SA_CollObjectNameList rl)
        {
            List<string> refList = new List<string>();
            foreach (SA_CollObjectName objname in rl.collObjNameRefList)
                refList.Add(objname.GetCollObjName());
            return DeleteObjects(ref refList);
        }

        private bool DeleteObjects(ref List<String> objsToDelete)
        {
            SetStep("Delete Objects");
            Object vObjList = null;
            MakeRefObjectList(objsToDelete, ref vObjList);
            SetCollectionObjectNameRefListArg("Object Names", ref vObjList);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool DeletePoints(SA_PointNameRefList pnrl)
        {
            SetStep("Delete Points");
            Object vObjList = null;
            MakePtNameObjectList(pnrl, ref vObjList);
            SetPointNameRefListArg("Point Names", ref vObjList);
            ExecuteStep();
            return IsDoneSuccess();
        }

        private void MakeRefObjectList(List<String> sLst, ref Object vObjList)
        {
            Object o = new ArrayList();
            foreach (String str in sLst)
                ((ArrayList)o).Add(str);
            vObjList = new VariantWrapper(((ArrayList)o).ToArray());
        }

        private void MakePtNameObjectList(SA_PointNameRefList pnrl, ref Object vObjList)
        {
            Object o = new ArrayList();
            foreach (SA_TargetName tn in pnrl.ptNameRefList)
                ((ArrayList)o).Add(tn.GetCollGrpTargetName());
            vObjList = new VariantWrapper(((ArrayList)o).ToArray());
        }


        private void MakeInstList(List<String> iLst, ref Object vInstList)
        {
            Object o = new ArrayList();
            foreach (String str in iLst)
                ((ArrayList)o).Add(str);
            vInstList = new VariantWrapper(((ArrayList)o).ToArray());
        }

        public bool MakeQuickReport(SA_CollObjectName vg, SA_Report r)
        {
            SetStep("Quick Report");
            SetCollectionObjectNameArg("Item Name", "", vg.GetObjName());
            SetStringArg("Report Name (optional)", r.GetReportName());
            SetBoolArg("Open Report?", true);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool MakeNewSAReport(SA_Report report)
        {
            SetStep("Make New SA Report");
            SetCollectionObjectNameArg("New SA Report Name", report.report.GetCollName(), report.report.GetObjName());
            SetCollectionObjectNameArg("SA Report Template (optional)", "", "");
            ExecuteStep();

            return IsDoneSuccess();
        }
        public bool AppendItemsToSAReport(SA_Report report, List<string> objList)
        {
            SetStep("Append Items to SA Report");
            SetCollectionObjectNameArg("Report Name", report.report.GetCollName(), report.report.GetObjName());

            SetCollectionObjectNameRefListArg("Items To Report", objList);

            SetBoolArg("Show Report?", false);
            SetBoolArg("Begin On New Page?", false);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool DeleteCollection(string coll)
        {
            SetStep("Delete Collection");
            SetCollectionNameArg("Name of Collection to Delete", coll);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool DeleteSAReport(SA_Report report)
        {
            SetStep("Delete SA Report");
            SetCollectionObjectNameArg("Report Name", report.report.GetCollName(), report.report.GetObjName());
            ExecuteStep();
            return IsDoneSuccess();
        }
        public bool ExportQuickReportToPDF(SA_Report r, string path, bool bEmbedded = false, bool bShowPDF = false)
        {
            SetStep("Output SA Report to PDF");
            SetCollectionObjectNameArg("Report Name", r.report.GetCollName(), r.GetReportName());
            SetFilePathArg("File Name", path, bEmbedded);
            SetBoolArg("Show PDF?", bShowPDF);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool GetVectorGroupProperties(VectorGroupProperties vgp, ToleranceSet ts)
        {

            SetStep("Get Vector Group Properties");
            SetCollectionObjectNameArg("Vector Group Name", "", vgp.VectorGroupName.GetObjName());
            bool ret = ExecuteStep();

            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetIntegerArg("Total Vectors", ref vgp.totVectors);
                    GetIntegerArg("Vectors In Tolerance", ref vgp.numVectorsInTol);
                    GetIntegerArg("Vectors Out Of Tolerance", ref vgp.numVectorsOutTol);
                    GetDoubleArg("% Vectors In Tolerance", ref vgp.percentVectorInTol);
                    GetDoubleArg("% Vectors Out Of Tolerance", ref vgp.percentVectorOutTol);
                    GetDoubleArg("Absolute Max Magnitude", ref vgp.absMaxMagResult);
                    GetDoubleArg("Absolute Min Magnitude", ref vgp.absMinMagResult);
                    GetDoubleArg("Max Magnitude", ref vgp.maxMag);
                    GetDoubleArg("Min Magnitude", ref vgp.minMag);
                    GetDoubleArg("Standard Deviation", ref ts.stdevResult);
                    GetDoubleArg("Standard Deviation Mean Zero", ref vgp.stdevMeanZero);
                    GetDoubleArg("Avg Magnitude", ref ts.avgResult);
                    GetDoubleArg("Avg of Abs Magnitude", ref vgp.avgAbsMag);
                    GetDoubleArg("High Tolerance Value", ref vgp.highTolValue);
                    GetDoubleArg("Low Tolerance Value", ref vgp.lowTolValue);
                    GetDoubleArg("RMS Value", ref ts.rmsResult);
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        private void SetCollectionGroupNameRefListArg(string arg, List<String> objList)
        {
            Object o = null;
            MakeRefObjectList(objList, ref o);
            SetCollectionGroupNameRefListArg(arg, ref o);
        }

        private void SetCollectionObjectNameRefListArg(string arg, List<String> objList)
        {
            Object o = null;
            MakeRefObjectList(objList, ref o);
            SetCollectionObjectNameRefListArg(arg, ref o);
        }

        private void SetCollectionInstrumentRefListArg(string arg, List<String> instList)
        {
            Object o = null;
            MakeInstList(instList, ref o);
            SetCollectionObjectNameRefListArg(arg, ref o);
        }

        public bool MakeGroupsToObjectsRelationship(SA_CollObjectName ptGrp, SA_CollObjectName obj)
        {
            List<String> ptGrpList = new List<String>();
            List<String> surfaceList = new List<String>();

            ptGrpList.Add(ptGrp.GetCollObjName());

            SA_Type t = new SA_Type();

            t.SetType(TypeIdx.eSurface);
            bool ret = MakeACollectionObjectNameRefList_ByType(obj.GetCollName(), t, ref surfaceList);
            if (ret)
            {
                SetStep("Make Groups to Objects Relationship");
                SetCollectionObjectNameArg("Relationship Name", "", "Rel" + ptGrp.GetObjName() + "_" + obj.GetObjName());
                SetCollectionObjectNameRefListArg("Point Groups in Relationship", ptGrpList);
                SetCollectionObjectNameRefListArg("Objects in Relationship", surfaceList);
                SetProjectionOptionsArg("Projection Options", "Object To Probe Vectors", false, false, 0.000000, false, 0.000000);
                SetBoolArg("Auto Update a Vector Group?", false);
                ExecuteStep();
                ret = IsDoneSuccess();
            }

            return ret;
        }

        public bool DoRelationshipFit(SA_CollObjectName ptGrp, ref Transform T)
        {
            List<String> ptGrpList = new List<String>();
            List<String> instIdList = new List<String>();

            ptGrpList.Add(ptGrp.GetCollObjName());

            bool ret = MakeACollectionInstrumentRefList("A", ref instIdList);

            if (!ret)
                return ret;

            SetStep("Do Relationship Fit");
            SetCollectionNameArg("Collection Containing Relationships", "");
            SetCollectionObjectNameRefListArg("Objects to Move", ptGrpList);
            SetCollectionInstrumentRefListArg("Instruments to Move", instIdList);
            SetBoolArg("Perform 'Direct' Search", false);
            SetFitDofOptionsArg("Motion to allow", true, true, true, true, true, true, true);
            SetBoolArg("Use Fit Dialog", false);

            ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    ret = GetTransformInWorkingFromSA(ref T);
                }
                else
                {
                    ret = false;
                }
            }

            return ret;
        }

        public bool GetSAVersion(ref string saVersion)
        {
            SetStep("Make a System String");
            // Available options: 
            // "SA Version", "XIT Filename", "MP Filename", "MP Filename (Full Path)", "Date & Time", 
            // "Date", "Date (Short)", "Time", 
            SetSystemStringArg("String Content", "SA Version");
            SetStringArg("Format String (Optional)", "");
            bool ret = ExecuteStep();

            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    ret = GetStringArg("Resultant String", ref saVersion);
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        public bool GetSAVersionDate(string saVersion, ref DateTime saVersionDate)
        {
            bool ret = false;
            try
            {
                int result = saVersion.IndexOf(" 20");
                if (result > 0)
                {
                    string yr = saVersion.Substring(result + 1);
                    int endOfyr = yr.IndexOf('.');
                    if (endOfyr > 0)
                    {
                        string sub = yr.Substring(0, endOfyr);
                        int y = Convert.ToInt32(sub);
                        sub = yr.Substring(endOfyr + 1, 2);
                        int m = Convert.ToInt32(sub);
                        sub = yr.Substring(endOfyr + 4, 2);
                        int d = Convert.ToInt32(sub);

                        saVersionDate = new DateTime(y, m, d);
                        ret = true;
                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("Could not read SA Version date from System String: " + saVersion);
                Trace.Write("Could not read SA Version date from System String: " + saVersion + " Ex: " + ex.ToString() + "/r/n");
            }

            return ret;
        }

        public bool SetCollectionObjectRefListVariable(String vName, List<String> objNameList)
        {
            SetStep("Set Collection Object Ref List Variable");
            SetStringArg("Name", vName);
            SetCollectionObjectNameRefListArg("Value", objNameList);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool SetCollectionObjectNameVariable(String vName, String objName)
        {
            SetStep("Set Collection Object Name Variable");
            SetStringArg("Name", vName);
            SetCollectionObjectNameArg("Value", "", objName);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool SetCollectionObjectNameVariable(String vName, SA_CollObjectName objName)
        {
            SetStep("Set Collection Object Name Variable");
            SetStringArg("Name", vName);
            SetCollectionObjectNameArg("Value", objName.GetCollName(), objName.GetObjName());
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool SetStringVariable(String vName, String str)
        {
            SetStep("Set String Variable");
            SetStringArg("Name", vName);
            SetStringArg("Value", str);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool SetDoubleVariable(String vName, double v)
        {
            SetStep("Set Double Variable");
            SetStringArg("Name", vName);
            SetDoubleArg("Value", v);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool GetDoubleVariable(String vName, ref double v)
        {
            SetStep("Get Double Variable");
            SetStringArg("Name", vName);
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetDoubleArg("Value", ref v);
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        public bool GetIntegerVariable(String vName, ref int v)
        {
            SetStep("Get Integer Variable");
            SetStringArg("Name", vName);
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetIntegerArg("Value", ref v);
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        public bool GetStringVariable(String vName, ref string v)
        {
            SetStep("Get String Variable");
            SetStringArg("Name", vName);
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetStringArg("Value", ref v);
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        public bool GetBooleanVariable(String vName, ref bool v)
        {
            SetStep("Get Boolean Variable");
            SetStringArg("Name", vName);
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetBoolArg("Value", ref v);
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        public bool GetLastObjectInCollection(SA_Type t, ref SA_CollObjectName obj)
        {
            List<string> sl = new List<string>();
            bool ret = MakeACollectionObjectNameRefList_ByType(obj.GetCollName(), t, ref sl);

            if (ret)
            {
                if (sl.Count > 0)
                {
                    obj.SetCollObjectNameType(sl.ElementAt(sl.Count - 1));
                }
            }
            return ret;
        }

        public bool ConstructSurfacesByDissectingSurfacesFromRefList(ref List<string> surfaceList)
        {
            Object vObjList = null;

            MakeRefObjectList(surfaceList, ref vObjList);
            SetStep("Construct Surfaces by Dissecting Surfaces from Ref List");
            SetCollectionObjectNameRefListArg("Surfaces to Dissect", surfaceList);

            bool ret = ExecuteStep();

            surfaceList.Clear();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    ret = GetCollectionObjectRefListFromSA("Resultant Surfaces List", ref surfaceList);
                }
                else
                {
                    ret = false;
                }
            }

            return ret;
        }

        public bool MakePointNameRefListRuntimeSelect(ref List<SA_TargetName> ptNameRefList)
        {
            SetStep("Make a Point Name Ref List - Runtime Select");
            SetStringArg("User Prompt", "Select Points to Construct Frames on:");
            bool ret = ExecuteStep();

            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetPointRefListFromSA("Resultant Point Name List", ref ptNameRefList);
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        public bool GetPointCoordinate(SA_TargetName tn, ref Vector3D v)
        {
            SetStep("Get Point Coordinate");
            SetPointNameArg("Point Name", tn.GetCollName(), tn.GetGrpName(), tn.GetTargetName());
            bool ret = ExecuteStep();

            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    ret = GetVectorArg("Vector Representation", ref v.x, ref v.y, ref v.z);
                }
                else
                    ret = false;
            }
            return ret;
        }

        public bool ConstructPointsFromSurfacesOnUV_Grid(List<String> surfaceList)
        {
            SetStep("Construct Points From Surfaces On UV Grid");
            SetCollectionGroupNameRefListArg("Surface List", surfaceList);
            SetStringArg("UV Point Group Base Name", "UV Points");
            SetBoolArg("Make Each Line Separate Group?", false);
            SetIntegerArg("Number of U Grids", 5);
            SetIntegerArg("Number of V Grids", 7);
            // Available options: 
            // "Include Edges", "Exclude Edges", "Edges Only", 
            SetEdgeModeArg("Edge Point Mode", "Exclude Edges");
            ExecuteStep();
            return IsDoneSuccess();
        }

        public void StripType(ref List<String> objList)
        {
            List<String> tempList = new List<String>();
            char[] delimiters = { ':' };
            foreach (String str in objList)
            {
                string[] words = str.Split(delimiters);
                String t = words[0] + "::" + words[2];
                tempList.Add(t);
            }
            objList = tempList;
        }

        public bool ConstructFrameFromPointMeasurementProbingFrames(List<string> pointNameList)
        {

            Object vObjList = null;

            MakeRefObjectList(pointNameList, ref vObjList);
            SetStep("Construct Frame From Point Measurement Probing Frames");
            SetPointNameRefListArg("Point List", ref vObjList);
            SetBoolArg("Show Frame? (Hide = FALSE)", false);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool SetInstrumentMeasurementModeProfile(SA_InstID ii, string measProfile)
        {
            SetStep("Set Instrument Measurement Mode/Profile");
            SetColInstIdArg("Instrument to set", ii.GetColl(), ii.GetInstIdx());
            SetStringArg("Mode/Profile", measProfile);
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    ret = true;
                }
                else
                {
                    ret = false;
                }
            }

            return ret;
        }

        public bool GetInstrumentMeasurementModeProfile(SA_InstID ii, ref List<string> measProfiles, ref List<string> targNames)
        {
            SetStep("Get Instrument Targets and Mode/Profiles");
            SetColInstIdArg("Instrument to set", ii.GetColl(), ii.GetInstIdx());
            bool ret = ExecuteStep();

            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    ret = GetStringRefListFromSA("Mode/Profile", ref measProfiles);
                    if (ret)
                        ret = GetStringRefListFromSA("Target Names", ref targNames);
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }


        public bool ConfigureAndMeasure(SA_TargetName tn, SA_InstID ii, string measurementMode, bool measImmediately, bool waitForCompletion, double timeout_seconds)
        {
            SetStep("Configure and Measure");
            SetColInstIdArg("Instrument's ID", ii.GetColl(), ii.GetInstIdx());
            SetPointNameArg("Target Name", tn.GetCollName(), tn.GetGrpName(), tn.GetTargetName());
            SetStringArg("Measurement Mode", measurementMode);
            SetBoolArg("Measure Immediately", measImmediately);
            SetBoolArg("Wait for Completion", waitForCompletion);
            SetDoubleArg("Timeout in Seconds", timeout_seconds);
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    ret = true;
                }
                else
                { 
                    ret = false; 
                }
            }
            return ret;

        }

        public bool MakeACollectionObjectNameReferenceListRuntimeSelect(string prompt, TypeIdx t, ref SA_CollObjectNameList rl)
        {
            SetStep("Make a Collection Object Name Reference List- Runtime Select");
            SetStringArg("User Prompt", prompt);
            // Available options: 
            // "Any", "B-Spline", "Circle", "Cloud", "Scan Stripe Cloud", 
            // "Cross Section Cloud", "Cone", "Cylinder", "Datum", "Ellipse", 
            // "Frame", "Frame Set", "Line", "Paraboloid", "Perimeter", 
            // "Plane", "Point Group", "Poly Surface", "Scan Stripe Mesh", "Slot", 
            // "Sphere", "Surface", "Vector Group", 
            SA_Type sa_type = new SA_Type(t);
            SetObjectTypeArg("Object Type", sa_type.GetTypeName());
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    List<string> rl_Names = new List<string>();
                    rl.GetStringList(ref rl_Names);
                    ret = GetCollectionObjectRefListFromSA("Resultant Collection Object Name Reference List", ref rl_Names);
                }
                else
                { 
                    ret = false; 
                }
            }

            return ret;
        }

        public bool PointAtTarget(SA_TargetName tn, SA_InstID ii)
        {
            SetStep("Point At Target");
            SetColInstIdArg("Instrument ID", ii.GetColl(), ii.GetInstIdx());
            SetPointNameArg("Target ID", tn.GetCollName(), tn.GetGrpName(), tn.GetTargetName());
            SetFilePathArg("HTML Prompt File (optional)", "", false);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool AddNewInstrument(ref SA_InstID inst)
        {
            SetStep("Add New Instrument");
            // Available options: 
            // "Faro Vantage", "Faro Tracker", "Faro Ion Tracker", "SMX Tracker 4000,4500", "Leica Tracker TP-LINK", 
            // "Leica emScon Tracker (LT500-800 Series)", "Leica emScon Absolute Tracker (AT901 Series)", "Leica emScon AT401", "Leica emScon AT402", "Leica emScon AT403", 
            // "Leica AT960/930", "API Tracker II", "API Tracker III", "API OmniTrac", "API Tracker Device Interface", 
            // "API Radian", "API OmniTrac2", "API Laser Rail", "Boeing Laser Tracker", "Chesapeake 3000 Laser Tracker", 
            // "Nikon Metrology Laser Radar MV200", "Nikon Metrology Laser Radar MV300", "Nikon Metrology Laser Radar (CLRICx)", "Nikon Metrology CLR 100 Laser Radar", "Boeing TaLLS Scanner", 
            // "Leica TPS Total Station (2003,5000,5005)", "Leica TDA5005 Total Station (GeoCOM)", "Leica Total Station TC2000, TC2002", "Leica Nova MS50 Total Station", "Leica Nova MS60 Total Station", 
            // "Leica TDRA6000 Total Station", "Sokkia SETX Total Station", "Sokkia Net05X Total Station", "Sokkia Net05AX Total Station", "Topcon MS AX Series Total Station", 
            // "Sokkia Net-1 Total Station", "Sokkia Net-2 Total Station", "FARO Arm", "FARO Arm G04", "FARO Arm S08", 
            // "FARO Arm G08", "FARO Arm S12", "FARO Arm G12", "FARO Arm G04-05 (7dof)", "FARO Arm G08-05 (7dof)", 
            // "FARO Arm G12-05 (7dof)", "FARO Arm USB 4 ft. (Quantum, Prime, Platinum)", "FARO Arm USB 6 ft. (Quantum, Fusion, Prime, Platinum)", "FARO Arm USB 8 ft.  (Quantum, Fusion, Prime, Platinum)", "FARO Arm USB 10 ft. (Quantum, Fusion, Prime, Platinum)", 
            // "FARO Arm USB 12 ft. (Quantum, Fusion, Prime, Platinum)", "FARO Arm USB 4 ft. 7 dof (Quantum, Prime, Platinum)", "FARO Arm USB 6 ft. 7 dof (Edge, Quantum, Fusion, Prime, Platinum)", "FARO Arm USB 8 ft. 7 dof (Quantum, Fusion, Prime, Platinum)", "FARO Arm USB 9 ft. 7 dof (Edge)", 
            // "FARO Arm USB 10 ft. 7 dof (Quantum, Fusion, Prime, Platinum)", "FARO Arm USB 12 ft. 7 dof (Edge, Quantum, Fusion, Prime, Platinum)", "CimCore Arm 1024", "CimCore Arm 1028", "CimCore Arm 1030", 
            // "CimCore Arm 2200", "CimCore Arm 2500", "CimCore Arm 6DOF: 3012i, 5012, 1.2m", "CimCore Arm 6DOF: 3018i, 5018, 1.8m", "CimCore Arm 6DOF: 3024i, 5024, 2.4m", 
            // "CimCore Arm 6DOF: 3028i, 5028, 2.8m", "CimCore Arm 6DOF: 3036i, 5036, 3.6m", "CimCore Arm 7DOF: 5012Sc, 3012, 1.2m", "CimCore Arm 7DOF: 5018Sc, 3018, 1.8m", "CimCore Arm 7DOF: 5030Sc, 3030, 3.0m", 
            // "CimCore Arm 7DOF: 5028Sc, 3028, 2.8m", "CimCore Arm 7DOF: 5024Sc, 3024, 2.4m", "CimCore Arm 7DOF: 5036Sc, 3036, 3.6m", "CimCore Arm 7DOF: 5112Sc, 1.2m", "CimCore Arm 7DOF: 5118Sc, 1.8m", 
            // "CimCore Arm 7DOF: 5130Sc, 3.0m", "CimCore Arm 7DOF: 5128Sc, 2.8m", "CimCore Arm 7DOF: 5124Sc, 2.4m", "CimCore Arm 7DOF: 5136Sc, 3.6m", "CimCore Arm 6DOF: 5112, 1.2m", 
            // "CimCore Arm 6DOF: 5118, 1.8m", "CimCore Arm 6DOF: 5130, 3.0m", "CimCore Arm 6DOF: 5128, 2.8m", "CimCore Arm 6DOF: 5124, 2.4m", "CimCore Arm 6DOF: 5136, 3.6m", 
            // "Romer Multi-Gage", "Romer Absolute 7x20SI/SE", "Romer Absolute 7x25SI/SE", "Romer Absolute 7x30SI/SE", "Romer Absolute 7x35SI/SE", 
            // "Romer Absolute 7x40SI/SE", "Romer Absolute 7x45SI/SE", "Romer Absolute 7315", "Romer Absolute 7x20", "Romer Absolute 7x25", 
            // "Romer Absolute 7x30", "Romer Absolute 7x35", "Romer Absolute 7x40", "Romer Absolute 7x45", "Nikon Metrology MCA Arm", 
            // "Romer Sigma Arm 2022", "Sandia National Labs Arm", "Axxis 6-100 Arm (2.6m 6 dof)", "Axxis 7-100 Arm Scanner (2.6m 7 dof)", "Axxis 7-100 Arm Probe (2.6m 7 dof)", 
            // "Axxis 6-200 Arm (3.2m 6 dof)", "Leica/Wild Theodolites T2000,T2002,T3000", "Leica TPS Theodolite (1800)", "Leica TPS Theodolite (5100)", "Zeiss ETh 2 Theodolite", 
            // "Kern E2 Theodolite", "Cubic KIT Theodolite", "GSI V-STARS Photogrammetry System", "AICON ProCam 3D Probe", "Nikon Metrology K-Series (K-Scan & SpaceProbe)", 
            // "METRONOR Portable Measurement System", "Creaform Handy Probe", "Nikon Metrology Surveyor", "Nikon Metrology iGPS Network", "Nikon Metrology iGPS Transmitter Simulator", 
            // "Nikon Metrology Surveyor v2", "Metron Scanner", "Minolta VIVID 700 Scanner", "Minolta VIVID 900 Scanner", "Nivel 20 Two Axis Level", 
            // "Thommen HM30 Weather Station", "On-Trak Laser Line System (OT-4040, OT-6000)", "Davis Perception II Weather Station", "ScAlert Temperature Probe", "Ultrasonic Thickness Gauge (CL400)", 
            // "Imported Measurements with Uncertainty", "Virtek Laser Projector", "LPT Laser Projector", "Assembly Guidance Laser Projector", "LAP CAD-Pro Laser Projector", 
            // "SA Open Instrument", "SA Open Auxiliary Instrument", "Faro Scanner Photon/LS/Focus 3D", "Surphaser Scanner", "Leica Geosystems ScanStation PXX", 
            // "Digital Network Level", "Creaform HandyScan 3D", "NDI OptoTrak", "Vicon Tracker", "AICON MoveInspect", 
            // "AICON DPA", "Leica TM6100A Theodolite", "Leica T1200 Total Station", "Leica TS15 Total Station", "Leica TS16 Total Station", 
            // "Leica TS30 Total Station", "Ubisense RTLS", "Creaform VXelements", "FARO Arm 2.5m 7 dof (QuantumS, QuantumM)", "FARO Arm 1.5m 6 dof (QuantumS, QuantumM)", 
            // "FARO Arm 3.5m 7 dof (QuantumS, QuantumM)", "FARO Arm 4m 7 dof (QuantumS, QuantumM)", "FARO Arm 2.5m 6 dof (QuantumS, QuantumM)", "FARO Arm 3.5m 6 dof (QuantumS, QuantumM)", "FARO Arm 4m 6 dof (QuantumS, QuantumM)", 
            SetInstTypeNameArg("Instrument Type", "Leica AT960/930");
            bool ret = ExecuteStep();

            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetColInstIdxFromSA(ref inst);
                }
                else
                { 
                    ret = false; 
                }
            }

            return ret;
        }

        public bool SetTargetComputationOptions(TargetComputationOptions targetComputationOption, bool bIgnoreDistanceMeasurements)
        { 
            SetStep("Set Target Computation Options");
	        // Available options: 
	        // "Use only most recent shot", "Use most recent shot from each face", "Do not change prior measurements at all", "Force a new point for each measurement", "Remove all prior shots", 
	        // "Deactivate all prior shots", 
	        SetTargetComputationMethodArg("Target Computation Method", targetComputationOption.GetTypeName());
	        SetBoolArg("Ignore Distance Measurements", bIgnoreDistanceMeasurements);
            ExecuteStep();
            return IsDoneSuccess();
        }


        public bool GetNumberOfCollection(ref int nColl)
        {
            SetStep("Get Number of Collections");
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetIntegerArg("Total Count", ref nColl);
                }
                else
                { 
                    ret = false; 
                }
            }
            return ret;
        }

        public bool Get_ithCollectionName(int nColl, ref string sCollName)
        {
            SetStep("Get i-th Collection Name");
            SetIntegerArg("Collection Index", nColl);
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    GetCollectionNameArg("Resultant Name", ref sCollName);
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }

        public bool GetNumberOfPointsInGroup(SA_CollObjectName ptgrp, ref int nPts)
        {
            SetStep("Get Number of Points in Group");
            SetCollectionObjectNameArg("Group Name", ptgrp.GetCollName(), ptgrp.GetObjName());
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    ret = GetIntegerArg("Total Count", ref nPts);
                }
                else
                { 
                    ret = false; 
                }
            }
            return ret;
        }
 


        public bool CreateRobotCalibration(string robCol, int robIdx, string calName)
        {
            SetStep("Create Robot Calibration");
            SetColMachineIdArg("Machine ID", robCol, robIdx);
            SetStringArg("Calibration Name", calName);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool DeleteRobotCalibration(string robCol, int robIdx, string calName)
        {
            SetStep("Delete Robot Calibration");
            SetColMachineIdArg("Machine ID", robCol, robIdx);
            SetStringArg("Calibration Name", calName);
            ExecuteStep();
            return IsDoneSuccess();
        }

        public bool GetObservationInfo(SA_TargetName tn, int oIdx, ref SA_InstID instIdx, ref Vector3D v, ref double rms)
        {
            SetStep("Get Observation Info");
            SetPointNameArg("Point Name", tn.GetCollName(), tn.GetGrpName(), tn.GetTargetName());
            SetIntegerArg("Observation Index", oIdx);
            bool ret = ExecuteStep();
            if (ret)
            {
                int rCode = 0;
                GetMPStepResult(ref rCode);
                if (rCode == (int)MPStatus.DoneSuccess)
                {
                    try
                    {
                        GetColInstIdArg("Resulting Instrument", ref instIdx.Coll, ref instIdx.instIdx);

                        GetVectorArg("Resultant Vector", ref v.x, ref v.y, ref v.z);
                        bool active = false;
                        GetBoolArg("Active?", ref active);

                        string timestamp = "";
                        GetStringArg("Timestamp", ref timestamp);
                        GetDoubleArg("RMS Error", ref rms);
                    }
                    catch (Exception ex)
                    {
                        Trace.Write("Trouble accessing target: " + tn.GetCollGrpTargetName() +
                            "'s observation info.  Execption: " + ex.ToString() + "\r\n");
                        ret = false;
                    }
                }
                else
                    ret = false;
            }

            return ret;
        }
    }
}
