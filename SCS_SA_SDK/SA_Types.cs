using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;
using System.Diagnostics;
using System.IO;

namespace SCS_SA_SDK
{
    [Serializable]
    public class Vector3D
    {
        public double x = 0.0;
        public double y = 0.0;
        public double z = 1.0;
        private int serializeVersion = 1;

        public Vector3D()
        {
            x = 0.0;
            y = 0.0;
            z = 1.0;
        }

        public Vector3D(double xV, double yV, double zV)
        {
            Set(xV, yV, zV);
        }

        public void Set(double xV, double yV, double zV)
        {
            x = xV;
            y = yV;
            z = zV;
        }

        public double Mag()
        {

            double dsquared = x * x + y * y + z * z;
            if (dsquared > 0)
                return Math.Sqrt(dsquared);
            else
                return 0.0;
        }
        //De serialization constructor.
        public Vector3D(SerializationInfo info, StreamingContext ctxt)
        {
            //Get the values from info and assign them to the appropriate properties

            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                x = info.GetDouble("x");
                y = info.GetDouble("y");
                z = info.GetDouble("z");
            }
        }
        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("x", x);
            info.AddValue("y", y);
            info.AddValue("z", z);

            // serializationVerion 2
        }

        public void FromSpherical(Vector3D rhv)
        {
            x = rhv.x * Math.Cos(rhv.y) * Math.Sin(rhv.z);
            y = rhv.x * Math.Sin(rhv.y) * Math.Sin(rhv.z);
            z = rhv.x * Math.Sin(rhv.z);
        }
        public string ToString(int sd, string delim)
        {
            string str = x.ToString("F" + sd) + delim + y.ToString("F" + sd) + delim + z.ToString("F" + sd);
            return str;
        }

        public static Vector3D operator +(Vector3D a, Vector3D b)
        {
            Vector3D v = NewMethod();
            v.x = a.x + b.x;
            v.y = a.y + b.y;
            v.z = a.z + b.z;
            return v;
        }

        private static Vector3D NewMethod()
        {
            return new Vector3D();
        }

        public static Vector3D operator -(Vector3D a)
        {
            Vector3D v = new Vector3D(-a.x, -a.y, -a.z);
            return v;
        }

        public static Vector3D operator -(Vector3D a, Vector3D b)
        {
            Vector3D v = new Vector3D();
            v.x = a.x - b.x;
            v.y = a.y - b.y;
            v.z = a.z - b.z;
            return v;
        }
        public void PositionReport(ref String pr, int sd)
        {
            Date_Time dt = new Date_Time();

            pr = "X:" + x.ToString("F" + sd) + ";Y:" + y.ToString("F" + sd) + ";Z:" + z.ToString("F" + sd) + dt.GetDateTime();
        }
        public void MagnitudeReport(ref String pr, int sd)
        {
            Date_Time dt = new Date_Time();

            pr = "Mag:" + Mag().ToString("F" + sd) + dt.GetDateTime();
        }

        public void Position_MagnitudeReport(ref String pr, int sd)
        {
            Date_Time dt = new Date_Time();

            pr = "X:" + x.ToString("F" + sd) + ";Y:" + y.ToString("F" + sd) + ";Z:" + z.ToString("F" + sd) + ";Mag:" + Mag().ToString("F" + sd) + dt.GetDateTime();
        }
    }

    [Serializable]
    public class ListVector3D
    {
        public List<Vector3D> listVectors = new List<Vector3D>(); 
        private int serializeVersion = 1;

        public ListVector3D()
        {
            listVectors.Clear();
        }

        public void Add(double xV, double yV, double zV)
        {
            Vector3D v = new Vector3D(xV, yV, zV);
            listVectors.Add(v);
        }

        public void Add(Vector3D v)
        {
            listVectors.Add(v);
        }

        public double RMS()
        {
           double rms = 0.0;
           if (listVectors.Count() > 0)
           {
                foreach (Vector3D v in listVectors)
                {
                    rms += v.Mag() * v.Mag();
                }
                rms = Math.Sqrt(rms / listVectors.Count());
            }
            return rms;
        }

        public double Mean()
        {
            double m = 0.0;
            if (listVectors.Count() > 0)
            {
                foreach (Vector3D v in listVectors)
                {
                    m += v.Mag();
                }
                m /= listVectors.Count();
            }
            return m;
        }

        public double Max()
        {
            double max = double.MinValue;
            if (listVectors.Count() > 0)
            {
                foreach (Vector3D v in listVectors)
                {
                    if (max < v.Mag())
                        max = v.Mag();
                }
            }
            else
                max = 0.0;
            return max;
        }

        public double Min()
        {
            double min = double.MaxValue;
            if (listVectors.Count() > 0)
            {
                foreach (Vector3D v in listVectors)
                {
                    if (min > v.Mag())
                        min = v.Mag();
                }
            }
            else
                min = 0.0;

            return min;
        }

        public string ResultsString(int sd)
        {
            string r = "Count: " + listVectors.Count() + 
                " RMS: " + RMS().ToString("F" + sd) + 
                " Max: " + Max().ToString("F" + sd);
            return r;
        }

        //De serialization constructor.
        public ListVector3D(SerializationInfo info, StreamingContext ctxt)
        {
            //Get the values from info and assign them to the appropriate properties
            //EmpId = (int)info.GetValue("EmployeeId", typeof(int));
            //EmpName = (String)info.GetValue("EmployeeName", typeof(string));

            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                int numCount = (int) info.GetValue("numVectors", typeof(int));
               
                for (int i = 0; i < numCount; i++)
                {
                    Vector3D v = new Vector3D(info, ctxt);
                    listVectors.Add(v);
                }
            }
        }
        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("numVectors", listVectors.Count());
            foreach (Vector3D v in listVectors)
                v.GetObjectData(info, ctxt);

            // serializationVerion 2
        }
    }

    [Serializable]
    public class Transform
    {
        public double[,] T = new double[4, 4];
        readonly double RadToDeg = 180.0 / Math.PI;
        Vector3D XYZ = new Vector3D();
        Vector3D rotXYZ = new Vector3D();
        
        private int serializeVersion = 1;

        public Transform()
        {
            SetIdentity();
        }

        public void SetIdentity()
        {
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    if (i == j)
                        T[i, j] = 1.0;
                    else
                        T[i, j] = 0.0;
                }
            }
        }


        public Transform Multiply(Transform t)
        {
            Transform m = new Transform();
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    m.T[i, j] = 0;
                    for (int k = 0; k < 4; k++)
                    {
                        m.T[i, j] += T[i, k] * t.T[k, j];
                    }
                }
            }
            return m;
        }

        public void TraceTransform()
        {
            Trace.Write("\n");
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    Trace.Write("T[" + i + "," + j +"] " + T[i,j].ToString("F" + 4) + " ");
                }
                Trace.Write("\n");
            }
        }

        //De serialization constructor.
        public Transform(SerializationInfo info, StreamingContext ctxt)
        {
            //Get the values from info and assign them to the appropriate properties

            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (serializeVersion >= 1) // serializationVerion 1
            {
                for (int i = 0; i < 4; i++)
                {
                    for (int j = 0; j < 4; j++)
                    {
                        String arrayIdxName = "T" + i.ToString() + j.ToString();
                        T[i, j] = info.GetDouble(arrayIdxName);
                    }
                }
                XYZ = new Vector3D(info, ctxt);
                rotXYZ = new Vector3D(info, ctxt);
            }
        }
        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 4; j++)
                { 
                    String arrayIdxName = "T" + i.ToString() + j.ToString();
                    info.AddValue(arrayIdxName, T[i, j]);
                }
                XYZ.GetObjectData(info, ctxt);
                rotXYZ.GetObjectData(info, ctxt);
            }
            // serializationVerion 2
        }

        public String TransformRowReport(int i, int sd)
        {
            String row = "Row " + (i + 1).ToString() + " " + T[i, 0].ToString("F" + sd) +
                                                 " " + T[i, 1].ToString("F" + sd) +
                                                 " " + T[i, 2].ToString("F" + sd) +
                                                 " " + T[i, 3].ToString("F" + sd);
            return row;
        }

        public void FixedAngleTransformRowReport(ref String Trans, ref String Rot, int sd)
        {
            FixedAnglesTransformFromMatrix();

            Trans = XYZ.x.ToString("F" + sd) + " " + XYZ.y.ToString("F" + sd) + " " + XYZ.z.ToString("F" + sd);
            Rot = rotXYZ.x.ToString("F" + sd) + " " + rotXYZ.y.ToString("F" + sd) + " " + rotXYZ.z.ToString("F" + sd) + " deg";
        }

        public void PositionReport(ref String pr, String name, int sd)
        {
            FixedAnglesTransformFromMatrix();
            Date_Time dt = new Date_Time();

            pr = name + ";X:" + XYZ.x.ToString("F" + sd) + ";Y:" + XYZ.y.ToString("F" + sd) + ";Z:" + XYZ.z.ToString("F" + sd) +
                ";Rx:" + rotXYZ.x.ToString("F" + sd) + ";Ry:" + rotXYZ.y.ToString("F" + sd) + ";Rz:" + rotXYZ.z.ToString("F" + sd) + 
                dt.GetDateTime();
        }

        private void FixedAnglesTransformFromMatrix()
        {
            XYZ.x = T[0, 3];
            XYZ.y = T[1, 3];
            XYZ.z = T[2, 3];

            rotXYZ.y = -Math.Asin(T[2, 0]);  // R(y) radians 
            double cosTheta = Math.Cos(rotXYZ.y);
            rotXYZ.y *= RadToDeg;  // R(y) degrees

            rotXYZ.x = 0;
            if (cosTheta != 0.0)
            {
                if (Math.Abs(T[2, 2]) > 0.0000001)
                    rotXYZ.x = Math.Atan2(T[2, 1] / cosTheta, T[2, 2] / cosTheta) * RadToDeg;
            }

            rotXYZ.z = 0;
            if (Math.Abs(T[0, 0]) > 0.0000001)
                rotXYZ.z = Math.Atan2(T[1, 0] / cosTheta, T[0, 0] / cosTheta) * RadToDeg;
        }

        public void FixedAltAzRotAnglesToTransform(Vector3D v)
        {
            Vector3D vR = new Vector3D(v.x / RadToDeg, v.y / RadToDeg, v.z / RadToDeg);
            T[0, 0] = Math.Cos(vR.y) * Math.Cos(vR.z);
            T[0, 1] = -Math.Cos(vR.x) * Math.Sin(vR.z) + Math.Sin(vR.x) * Math.Sin(vR.y) * Math.Cos(vR.z);
            T[0, 2] = Math.Sin(vR.x) * Math.Sin(vR.z) + Math.Cos(vR.x) * Math.Sin(vR.z) * Math.Cos(vR.z);
            T[1, 0] = Math.Cos(vR.y) * Math.Sin(vR.z);
            T[1, 1] = Math.Cos(vR.x) * Math.Cos(vR.z) + Math.Sin(vR.x) * Math.Sin(vR.y) * Math.Sin(vR.z);
            T[1, 2] = -Math.Sin(vR.x) * Math.Cos(vR.z) + Math.Cos(vR.x) * Math.Sin(vR.z) * Math.Sin(vR.z);
            T[3, 0] = -Math.Sin(vR.y);
            T[3, 1] = Math.Sin(vR.x) * Math.Cos(vR.y);
            T[3, 2] = Math.Cos(vR.x) * Math.Cos(vR.y);
        }
    }

    public class Date_Time
    {
        public string GetDateTime()
        {
            DateTime dt = DateTime.Now;
            return dt.ToString(";MM/dd/yyyy HH:mm:ss");
        }

        public void AddDateTimeString(ref string s)
        {
            s += GetDateTime();
        }
    }

    [Serializable]
    public class SA_Units
    {
        public String lengthUnit = ""; // "Meters", "Centimeters", "Millimeters", "Feet", "Inches", "" leave blank equals native CAD file units
        public String angUnit = "Degrees"; // "Degrees");
        public List<string> listLengthUnits = new List<string>();
        public List<string> listAngleUnits = new List<string>();

        private int serializeVersion = 1;

        public void SetLengthUnits(String unit)
        {
            lengthUnit = unit;
        }

        public void SetAngularUnits(String unit)
        {
            angUnit = unit;
        }

        public SA_Units(String lenUnit, String angUnit)
        {
            Load_unitTypes();
            SetLengthUnits(lenUnit);
            SetAngularUnits(angUnit);
        }

        public SA_Units()
        {
            Load_unitTypes();
        }


        private void Load_unitTypes()
        {
            listLengthUnits.Add("Meters");
            listLengthUnits.Add("Centimeters");
            listLengthUnits.Add("Millimeters");
            listLengthUnits.Add("Feet");
            listLengthUnits.Add("Inches");
            listLengthUnits.Add(""); // blank equal native CAD file units

            listAngleUnits.Add("Degrees");
            listAngleUnits.Add("Deg:Min:Sec");
            listAngleUnits.Add("Radians");
            listAngleUnits.Add("Milliradians");
            listAngleUnits.Add("Gons/Grad");
            listAngleUnits.Add("Mils");
            listAngleUnits.Add("Arcseconds");
            listAngleUnits.Add("Deg:Min");

        }
        //De serialization constructor.
        public SA_Units(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                lengthUnit = info.GetString("lengthUnit");
                angUnit = info.GetString("angUnit");
            }
        }

        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("lengthUnit", lengthUnit);
            info.AddValue("angUnit", angUnit);

            // serializationVerion 2
        }

        public String GetLengthUnit()
        {
            return lengthUnit;
        }

        public int GetLengthUnitIdx(String u)
        {
            int idx = 0;
            foreach (String s in listLengthUnits)
            {
                if (s.Equals(u))
                    return idx;
                idx++;
            }
            return 5; // index for empty
        }

        public String GetAngleUnit()
        {
            return angUnit;
        }

        public int GetActiveLengthUnitIdx()
        {
            int idx = 0;
            foreach (String s in listLengthUnits)
            {
                if (s.Equals(lengthUnit))
                    return idx;
                idx++;
            }
            return 5; // index for empty as a default
        }

        public int GetAngleUnitIdx(String u)
        {
            int idx = 0;
            foreach (String s in listAngleUnits)
            {
                if (s.Equals(u))
                    return idx;
                idx++;
            }
            return 0; // index for degrees
        }

        public int GetActiveAngleUnitIdx()
        {
            int idx = 0;
            foreach (String s in listAngleUnits)
            {
                if (s.Equals(angUnit))
                    return idx;
                idx++;
            }
            return 0; // index for degrees
        }
    }

    [Serializable]
    public class SA_ASCIIFileFormats
    {
        public String ASCIIFileFormat = "PointName X Y Z";
        public int formatIdx = 5; // index for PointName X Y Z
        public List<string> listASCIIFileFormats = new List<string>();

        private int serializeVersion = 1;

        public void SetFileFormat(String format)
        {
            ASCIIFileFormat = format;
        }

        public SA_ASCIIFileFormats()
        {
            // Available options: 
            listASCIIFileFormats.Add("X Y Z");
            listASCIIFileFormats.Add("X Y Z Offset [Offset2]");
            listASCIIFileFormats.Add("X Y Z [Notes]");
            listASCIIFileFormats.Add("Radius Theta Phi (polar or spheric)");
            listASCIIFileFormats.Add("Radius Theta Z (cylindric)");
            listASCIIFileFormats.Add("PointName X Y Z");
            listASCIIFileFormats.Add("PointName X Y Z [Notes]");
            listASCIIFileFormats.Add("PointName X Y Z Offset [Offset2]");
            listASCIIFileFormats.Add("PointName X Y Z Ux Uy Uz (1 sigma)");
            listASCIIFileFormats.Add("PointName X Y Z Tx Ty Tz Td (Point Tolerance)");
            listASCIIFileFormats.Add("PointName X Y Z Wx Wy Wz [Wmag]");
            listASCIIFileFormats.Add("PointName X Y Z THx TLx THy TLy THz TLz THd TLd (Point Tolerance)");
            listASCIIFileFormats.Add("PointName X Y Z Tx Ty Tz Td Wx Wy Wz");
            listASCIIFileFormats.Add("PointName X Y Z Wx Wy Wz Tx Ty Tz Td");
            listASCIIFileFormats.Add("PointName X Y Z THx TLx THy TLy THz TLz THd TLd Wx Wy Wz");
            listASCIIFileFormats.Add("PointName X Y Z Wx Wy Wz THx TLx THy TLy THz TLz THd TLd");
            listASCIIFileFormats.Add("PointName Radius Theta Phi (polar or spheric)");
            listASCIIFileFormats.Add("PointName Radius Theta Z (cylindric)");
            listASCIIFileFormats.Add("PointName X Y Z GroupName");
            listASCIIFileFormats.Add("PointName Y X Z GroupName");
            listASCIIFileFormats.Add("GroupName PointName X Y Z");
            listASCIIFileFormats.Add("GroupName PointName X Y Z Offset [Offset2]");
            listASCIIFileFormats.Add("GroupName PointName X Y Z [Notes]");
            listASCIIFileFormats.Add("GroupName PointName X Y Z Ux Uy Uz (1 sigma)");
            listASCIIFileFormats.Add("GroupName PointName Radius Theta Phi");
            listASCIIFileFormats.Add("GroupName PointName Radius Theta Z");
            listASCIIFileFormats.Add("Collection Group Point X Y Z");
            listASCIIFileFormats.Add("Collection Group Point X Y Z [Notes]");
            listASCIIFileFormats.Add("Collection Group Point Radius Theta Phi");
            listASCIIFileFormats.Add("Collection Group Point Radius Theta Z");
            listASCIIFileFormats.Add("X Y Z I J K (Planes or Vectors)");
            listASCIIFileFormats.Add("VectorName X Y Z I J K");
            listASCIIFileFormats.Add("VectorName X Y Z dX dY dZ SignedMag(optional)");
            listASCIIFileFormats.Add("VectorGroupName VectorName X Y Z I J K");
            listASCIIFileFormats.Add("VectorGroupName VectorName X Y Z dX dY dZ SignedMag(optional)");
            listASCIIFileFormats.Add("FrameName X Y Z  Rx Ry Rz");
            listASCIIFileFormats.Add("FrameName X Y Z  Euler XYZ");
            listASCIIFileFormats.Add("FrameName X Y Z  Euler ZYX");
            listASCIIFileFormats.Add("FrameName X Y Z  Euler ZYZ");
            listASCIIFileFormats.Add("FrameName X Y Z  Euler ZXZ");
            listASCIIFileFormats.Add("PlaneName X Y Z dX dY dZ PlaneSize(optional)");
        }

        //De serialization constructor.
        public SA_ASCIIFileFormats(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                formatIdx = (int)info.GetValue("formatIdx", typeof(int));
                ASCIIFileFormat = info.GetString("ASCIIFileFormat");
            }
        }

        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("formatIdx", formatIdx);
            info.AddValue("ASCIIFileFormat", ASCIIFileFormat);
            // serializationVerion 2
        }
    }

    [Serializable]
    public class ASCIIPredefinedFormatOptions
    {
        private int serializeVersion = 1;
        public SA_ASCIIFileFormats asciiFileFormat;
        public bool embeddedFile = false;
        public bool bImportWarnings = false;
        public String sImportWarningMessages = "";

        // Available options: 
        // "X Y Z", "X Y Z Offset [Offset2]", "X Y Z [Notes]", "Radius Theta Phi (polar or spheric)", "Radius Theta Z (cylindric)", 
        // "PointName X Y Z", "PointName X Y Z [Notes]", "PointName X Y Z Offset [Offset2]", "PointName X Y Z Ux Uy Uz (1 sigma)", "PointName X Y Z Tx Ty Tz Td (Point Tolerance)", 
        // "PointName X Y Z Wx Wy Wz [Wmag]", "PointName X Y Z THx TLx THy TLy THz TLz THd TLd (Point Tolerance)", "PointName X Y Z Tx Ty Tz Td Wx Wy Wz", "PointName X Y Z Wx Wy Wz Tx Ty Tz Td", "PointName X Y Z THx TLx THy TLy THz TLz THd TLd Wx Wy Wz", 
        // "PointName X Y Z Wx Wy Wz THx TLx THy TLy THz TLz THd TLd", "PointName Radius Theta Phi (polar or spheric)", "PointName Radius Theta Z (cylindric)", "PointName X Y Z GroupName", "PointName Y X Z GroupName", 
        // "GroupName PointName X Y Z", "GroupName PointName X Y Z Offset [Offset2]", "GroupName PointName X Y Z [Notes]", "GroupName PointName X Y Z Ux Uy Uz (1 sigma)", "GroupName PointName Radius Theta Phi", 
        // "GroupName PointName Radius Theta Z", "Collection Group Point X Y Z", "Collection Group Point X Y Z [Notes]", "Collection Group Point Radius Theta Phi", "Collection Group Point Radius Theta Z", 
        // "X Y Z I J K (Planes or Vectors)", "VectorName X Y Z I J K", "VectorName X Y Z dX dY dZ SignedMag(optional)", "VectorGroupName VectorName X Y Z I J K", "VectorGroupName VectorName X Y Z dX dY dZ SignedMag(optional)", 
        // "FrameName X Y Z  Rx Ry Rz", "FrameName X Y Z  Euler XYZ", "FrameName X Y Z  Euler ZYX", "FrameName X Y Z  Euler ZYZ", "FrameName X Y Z  Euler ZXZ", 
        // "PlaneName X Y Z dX dY dZ PlaneSize(optional)", 
        // public String setAsciiFileFormat = @"PointName X Y Z"; // ("File Format", "");
        // Available options: 
        // "Meters", "Centimeters", "Millimeters", "Feet", "Inches", 
        public SA_Units units; // public String SetDistanceUnitsArg("Units", "Millimeters");
        // Available options: 
        // "Degrees", "Deg:Min:Sec", "Radians", "Milliradians", "Gons/Grad", "Mils", "Arcseconds", "Deg:Min", 
        // public String SetAngularUnitsArg("Angular Units", "Degrees");
        public SA_CollObjectName grpName; //SetCollectionObjectNameArg("Group Name", "", "");
        public bool importAsCloud = false;

        public ASCIIPredefinedFormatOptions()
        {
            units = new SA_Units("Inches", "Degrees");
            grpName = new SA_CollObjectName("", "RefPtGrp", TypeIdx.ePointGroup);
            asciiFileFormat = new SA_ASCIIFileFormats();
        }

        public ASCIIPredefinedFormatOptions(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                embeddedFile = info.GetBoolean("embeddedFile");
                importAsCloud = info.GetBoolean("importAsCloud");
                bImportWarnings = info.GetBoolean("bImportWarning");
                sImportWarningMessages = info.GetString("sImportWarningMessages");

                units = new SA_Units(info, ctxt);
                grpName = new SA_CollObjectName(info, ctxt);

                asciiFileFormat = new SA_ASCIIFileFormats(info, ctxt);
            }
        }
        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("embeddedFile", embeddedFile);

            info.AddValue("importAsCloud", importAsCloud);
            info.AddValue("bImportWarning", bImportWarnings);
            info.AddValue("sImportWarningMessages", sImportWarningMessages);

            units.GetObjectData(info, ctxt);
            grpName.GetObjectData(info, ctxt);
            asciiFileFormat.GetObjectData(info, ctxt);

            // serializationVerion 2
        }
    }
    public enum TypeIdx
    {
        eAny = 0,
        eBSpline = 1,
        eCircle = 2,
        eCloud = 3,
        eScanStripCloud = 4,
        eCrossSectionCloud = 5,
        eCone = 6,
        eCylinder = 7,
        eDatum = 8,
        eEllipse = 9,
        eFrame = 10,
        eFrameSet = 11,
        eLine = 12,
        eParaboloid = 13,
        ePerimeter = 14,
        ePlane = 15,
        ePointGroup = 16,
        ePolySurface = 17,
        eScanStripeMesh = 18,
        eSlot = 19,
        eSphere = 20,
        eSurface = 21,
        eVectorGroup = 22,
        eSAReport = 23
    };

    [Serializable]
    public class SA_Type
    {
        public TypeIdx typeIdx = TypeIdx.eAny;

        public String[] typeNames = {"Any", "B-Spline", "Circle", "Cloud", "Scan Stripe Cloud",
                                    "Cross Section Cloud", "Cone", "Cylinder",  "Datum",
                                    "Ellipse", "Frame", "Frame Set", "Line", "Paraboloid", "Perimeter", 
                                    "Plane", "Point Group", "Poly Surface",  "Scan Stripe Mesh", "Slot", 
                                    "Sphere", "Surface", "Vector Group"};
	// Available options: 
	// "Any", "B-Spline", "Circle", "Cloud", "Scan Stripe Cloud", 
	// "Cross Section Cloud", "Cone", "Cylinder", "Datum", "Ellipse", 
	// "Frame", "Frame Set", "Line", "Paraboloid", "Perimeter", 
	// "Plane", "Point Group", "Poly Surface", "Scan Stripe Mesh", "Slot", 
	// "Sphere", "Surface", "Vector Group", 

        int serializeVersion = 1;
        public SA_Type()
        {
            typeIdx = TypeIdx.eAny;
        }

        public SA_Type(TypeIdx idx)
        {
            typeIdx = (TypeIdx)idx;
        }

        public void SetType(TypeIdx idx)
        {
            typeIdx = idx;
        }

        public void SetType(string name)
        {
            int idx = 0;
            typeIdx = TypeIdx.eAny;
            foreach (string typename in typeNames)
            {
                idx++;
                if (typename == name)
                {
                    typeIdx = (TypeIdx)idx;
                    break;
                }
            }
        }

        public String GetTypeName()
        {
            return typeNames.ElementAt((int)typeIdx);
        }

        public String GetTypeName(TypeIdx idx)
        {
            return typeNames.ElementAt((int)idx);
        }

        public SA_Type(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                typeIdx = (TypeIdx)info.GetValue("typeIdx", typeof(TypeIdx));
            }
        }

        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("typeIdx", typeIdx);

            // serializationVerion 2
        }

    }
    // Available Target Computation Method Options: 
    //
    public enum eTargetComputationOption
    {
        eUseOnlyMostRecentShot = 0,
        eUseMostRecentShotFromEachFace = 1,
        eDoNotChangePriorMeasurementsAtAll = 2,
        eForceANewPointForEachMeasurement = 3,
        eRemoveAllPriorShots = 4,
        eDeactiviateAllPriorShots = 5
    };

    [Serializable]
    public class TargetComputationOptions
    {
        public eTargetComputationOption targetCompOptionIdx = eTargetComputationOption.eUseOnlyMostRecentShot;
        public string[] typeNames = { "Use only most recent shot", 
                                     "Use most recent shot from each face", 
                                     "Do not change prior measurements at all", 
                                     "Force a new point for each measurement", 
                                     "Remove all prior shots", 
                                    "Deactivate all prior shots"
                                    };
        int serializeVersion = 1;

        public TargetComputationOptions()
        {
            targetCompOptionIdx = eTargetComputationOption.eUseOnlyMostRecentShot;
        }

        public TargetComputationOptions(eTargetComputationOption idx)
        {
            targetCompOptionIdx = idx;
        }

        public void SetType(eTargetComputationOption idx)
        {
            targetCompOptionIdx = idx;
        }

        public void SetType(string name)
        {
            int idx = 0;
            foreach (string typename in typeNames)
            {
                idx++;
                if (typename == name)
                {
                    targetCompOptionIdx = (eTargetComputationOption)idx;
                    break;
                }
            }
        }

        public string GetTypeName()
        {
            return typeNames.ElementAt((int)targetCompOptionIdx);
        }

        public string GetTypeName(eTargetComputationOption idx)
        {
            return typeNames.ElementAt((int)idx);
        }

        public TargetComputationOptions(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                targetCompOptionIdx = (eTargetComputationOption)info.GetValue("targetCompOptionIdx", typeof(eTargetComputationOption));
            }
        }

        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("targetCompOptionIdx", targetCompOptionIdx);

            // serializationVerion 2
        }

    }

    public enum DoFType
    {
        eX = 0,
        eY = 1,
        eZ = 2,
        eRx = 3,
        eRy = 4,
        eRz = 5,
        eScale = 6,
    };

    [Serializable]
    public class DoF
    {
        public bool dofX = true;
        public bool dofY = true;
        public bool dofZ = true;
        public bool dofRx = true;
        public bool dofRy = true;
        public bool dofRz = true;
        public bool dofScale = false;

        int serializeVersion = 1;

        public DoF()
        {
            dofX = true;
            dofY = true;
            dofZ = true;
            dofRx = true;
            dofRy = true;
            dofRz = true;
            dofScale = false;
        }
        public DoF(bool x, bool y, bool z, bool rx, bool ry, bool rz, bool scale = false)
        {
            dofX = x;
            dofY = y;
            dofZ = z;
            dofRx = rx;
            dofRy = ry;
            dofRz = rz;
            dofScale = scale;
        }

        public void SetDoF(bool x, bool y, bool z, bool rx, bool ry, bool rz, bool scale = false)
        {
            dofX = x;
            dofY = y;
            dofZ = z;
            dofRx = rx;
            dofRy = ry;
            dofRz = rz;
            dofScale = scale;
        }


        public DoF(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                dofX = (bool)info.GetValue("dofX", typeof(bool));
                dofY = (bool)info.GetValue("dofY", typeof(bool));
                dofZ = (bool)info.GetValue("dofZ", typeof(bool));
                dofRx = (bool)info.GetValue("dofRx", typeof(bool)); 
                dofRy = (bool)info.GetValue("dofRy", typeof(bool)); 
                dofRz = (bool)info.GetValue("dofRz", typeof(bool)); 
                dofScale = (bool)info.GetValue("scale", typeof(bool)); 
            }
        }

        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("dofX", dofX);
            info.AddValue("dofY", dofY);
            info.AddValue("dofZ", dofZ);
            info.AddValue("dofRx", dofRx);
            info.AddValue("dofRy", dofRy);
            info.AddValue("dofRz", dofRz);
            info.AddValue("dofScale", dofScale);

            // serializationVerion 2
        }

    }

    [Serializable]
    public class SA_CollObjectName
    {
        public String Coll = @"A"; 
        public String objectName = @"";

        public SA_Type saType = new SA_Type();

        int serializeVersion = 1;


        public SA_CollObjectName()
        {
            // typeName = typeNames.ElementAt(TypeIdx.eAny);
        }


        public SA_CollObjectName(String c, String objName, TypeIdx nameIdx)
        {
            Coll = c;
            objectName = objName;
            saType.SetType(nameIdx);
        }

        public void Clear()
        {
            Coll = String.Empty;
            objectName = String.Empty;
            saType = new SA_Type();
        }

        public String GetCollObjNameType()
        {
            String c_ObjName = Coll + "::" + objectName + "::" + saType.GetTypeName();
            return c_ObjName;
        }

        public void SetCollObjectName_Type(String c, String objName, TypeIdx t)
        {
            Coll = c;
            objectName = objName;
            saType.typeIdx = t;
        }
        
        public void SetCollObjectNameType(string con)
        {
            char[] delimiters = { ':', ',' };
            string[] words = con.Split(delimiters);
            if (words.Count() >= 4)
            {
                Coll = words[0];
                objectName = words[2];
                saType.SetType(words[4]);
            }
        }

        public String GetCollObjName()
        {
            String c_ObjName = Coll + "::" + objectName;
            return c_ObjName;
        }

        public override string ToString()
        {
            return GetCollObjName();
        }

        public String GetCollName()
        {
            return Coll;
        }
        public String GetObjName()
        {
            return objectName;
        }
        public String GetObjType()
        {
            return saType.GetTypeName();
        }

        public SA_CollObjectName(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                Coll = info.GetString("Coll");
                objectName = info.GetString("objectName");
                saType = new SA_Type(info, ctxt);
            }
        }

        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("Coll", Coll);
            info.AddValue("objectName", objectName);

            saType.GetObjectData(info, ctxt);

            // serializationVerion 2
        }
    }


    [Serializable]
    public class SA_CollObjectNameList
    {
        public List<SA_CollObjectName> collObjNameRefList = new List<SA_CollObjectName>();
        
        int serializeVersion = 1;


        public SA_CollObjectNameList()
        {
            // typeName = typeNames.ElementAt(TypeIdx.eAny);
        }


        public SA_CollObjectNameList(String c, String objName, TypeIdx nameIdx)
        {
            SA_CollObjectName con = new SA_CollObjectName(c, objName, nameIdx);
            collObjNameRefList.Add(con);
        }

        public String GetCollObjNameType(int idx)
        {
            string type = "";
            if (collObjNameRefList.Count <= idx && idx >= 0)
            {
                type = collObjNameRefList[idx].saType.GetTypeName();
            }
            return type;
        }

        public String GetCollObjNameIdx(int idx)
        {
            String c_ObjName = "";
            if (collObjNameRefList.Count <= idx && idx >= 0)
            {
                c_ObjName = collObjNameRefList[idx].GetCollObjNameType();
            }
            return c_ObjName;
        }

        public SA_CollObjectName GetCollObjName(int idx)
        {
            SA_CollObjectName c_ObjName = new SA_CollObjectName();
            if (collObjNameRefList.Count <= idx && idx >= 0)
            {
                c_ObjName = collObjNameRefList[idx];
            }
            return c_ObjName;
        }

        public bool RemoveCollObjectNameIdx(int idx)
        {
            bool ret = false;
            if (collObjNameRefList.Count <= idx && idx >= 0)
            {
                collObjNameRefList.RemoveAt(idx);
                ret = true;
            }
            return ret;

        }

        public void GetStringList(ref List<string> sl)
        {
            foreach (SA_CollObjectName con in collObjNameRefList)
            {
                sl.Add(con.GetCollObjName());
            }
        }

        public SA_CollObjectNameList(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                int count = (int)info.GetValue("count",typeof(int));
                for (int i = 0; i < count; i++)
                {
                    SA_CollObjectName con = new SA_CollObjectName(info, ctxt);
                    collObjNameRefList.Add(con);
                }
            }
        }

        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("count", collObjNameRefList.Count);
            foreach (SA_CollObjectName con in collObjNameRefList)
            {
                con.GetObjectData(info, ctxt);
            }

            // serializationVerion 2
        }
    }


    [Serializable]
    public class SA_TargetName
    {
        public String Coll = @"A";
        public String grpName = @"";
        public String targetName = @"";

        int serializeVersion = 2;
        /// <summary>
        /// /////// serialization = 2 added timestamp
        /// </summary>
        public DateTime timestamp; 

        public SA_TargetName()
        {
            // typeName = typeNames.ElementAt(TypeIdx.eAny);
            timestamp = DateTime.Now;
        }

        public SA_TargetName(string coll_grp_t)
        {
            char[] delimiters = { ':', ',' };
            string[] words = coll_grp_t.Split(delimiters);
            if (words.Count() >= 4)
            {
                Coll = words[0];
                grpName = words[2];
                targetName = words[4];
            }
        }


        public SA_TargetName(String c, String ptgrpName, String targName)
        {
            Coll = c;
            grpName = ptgrpName;
            targetName = targName;

            timestamp = DateTime.Now;
        }

        public String GetTargetName()
        {
            return targetName;
        }
        public String GetCollName()
        {
            return Coll; 
        }

        public String GetCollGrpName()
        {
            return Coll + "::" + grpName;
        }

        public String GetCollGrpTargetName()
        {
            return Coll + "::" + grpName + "::" + targetName;
        }

        public String GetGrpName()
        {
            return grpName;
        }

        public SA_TargetName(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                Coll = info.GetString("Coll");
                grpName = info.GetString("grpName");
                targetName = info.GetString("targetName");
            }
            if (sVersion >= 2)
            {
                timestamp = info.GetDateTime("timestamp");
            }
        }

        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("Coll", Coll);
            info.AddValue("grpName", grpName);
            info.AddValue("targetName", targetName);

            // serializationVerion 2
            info.AddValue("timestamp", timestamp);
        }
    }

    [Serializable]
    public class SA_PointNameRefList
    {
        public List<SA_TargetName> ptNameRefList = new List<SA_TargetName>();

        int serializeVersion = 1;


        public SA_PointNameRefList()
        {
        }

        public String GetPointNameAtIdx(int idx)
        {
            String c_PtName = "";
            if (ptNameRefList.Count <= idx && idx >= 0)
            {
                c_PtName = ptNameRefList[idx].GetTargetName();
            }
            return c_PtName;
        }

        public SA_TargetName GetCollGrpPtName(int idx)
        {
            SA_TargetName c_PtName = new SA_TargetName();
            if (ptNameRefList.Count <= idx && idx >= 0)
            {
                c_PtName = ptNameRefList[idx];
            }
            return c_PtName;
        }

        public bool RemovePtNameIdx(int idx)
        {
            bool ret = false;
            if (ptNameRefList.Count <= idx && idx >= 0)
            {
                ptNameRefList.RemoveAt(idx);
                ret = true;
            }
            return ret;

        }

        public void MakePointNameList(ref List <string> ptNameList)
        {
            foreach (SA_TargetName tn in ptNameRefList)
            {
                ptNameList.Add(tn.GetCollGrpTargetName());
            }
        }

        public void GetTargetNameStringList(ref List<string> sl)
        {
            foreach (SA_TargetName con in ptNameRefList)
            {
                sl.Add(con.GetTargetName());
            }
        }

        public void GetCollGrpPointNameStringList(ref List<string> sl)
        {
            foreach (SA_TargetName con in ptNameRefList)
            {
                sl.Add(con.GetCollGrpTargetName());
            }
        }

        public void AddSampleSets(int iterations)
        {
            int count = ptNameRefList.Count;
            for (int i = 1; i < iterations; i++)
            {
                for (int j = 0; j < count; j++)
                {
                    ptNameRefList.Add(ptNameRefList[j]);
                }
            }
        }
        public void Randomize()
        {
            List<SA_TargetName> temp = new List<SA_TargetName>();
            var random = new Random();

            while (ptNameRefList.Count > 0)
            {
                int idx = random.Next(ptNameRefList.Count);
                // Trace.WriteLine("Random: Count = " + ptNameRefList.Count.ToString() + " idx: " + idx.ToString() + " ptName: " + ptNameRefList[idx].GetCollGrpTargetName() + "\n");
                temp.Add(ptNameRefList[idx]);
                ptNameRefList.RemoveAt(idx);
            }
            ptNameRefList = temp;
        }
        public SA_PointNameRefList(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                int count = (int)info.GetValue("count",typeof(int));
                for (int i = 0; i < count; i++)
                {
                    SA_TargetName con = new SA_TargetName(info, ctxt);
                    ptNameRefList.Add(con);
                }
            }
        }

        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("count", ptNameRefList.Count);
            foreach (SA_TargetName con in ptNameRefList)
            {
                con.GetObjectData(info, ctxt);
            }

            // serializationVerion 2
        }
    }


    [Serializable]
    public class ToleranceSet
    {
        public double rmsTol = 0.0;
        public double maxTol = 0.0;
        public double avgTol = 0.0;
        public double rmsResult = 0.0;
        public double maxResult = 0.0;
        public double avgResult = 0.0;
        public double stdevTol = 0.0;
        public double stdevResult = 0.0;

        int serializeVersion = 1;

        public ToleranceSet()
        {
            ResetToleranceSet();
        }

        public void ResetToleranceSet()
        {
            rmsTol = 0.0;
            maxTol = 0.0;
            avgTol = 0.0;
            rmsResult = 0.0;
            maxResult = 0.0;
            avgResult = 0.0;
            stdevTol = 0.0;
            stdevResult = 0.0;
    }

        public ToleranceSet(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                rmsTol = info.GetDouble("rmsTol");
                rmsResult = info.GetDouble("rmsResult");
                maxTol = info.GetDouble("maxTol");
                maxResult = info.GetDouble("maxResult");
                avgTol = info.GetDouble("avgTol");
                avgResult = info.GetDouble("avgResult");
                stdevTol = info.GetDouble("stdevTol");
                stdevResult = info.GetDouble("stdevResult");
            }
        }

        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("rmsTol", rmsTol);
            info.AddValue("rmsResult", rmsResult);
            info.AddValue("maxTol", maxTol);
            info.AddValue("maxResult", maxResult);
            info.AddValue("avgTol", avgTol);
            info.AddValue("avgResult", avgResult);
            info.AddValue("stdevTol", stdevTol);
            info.AddValue("stdevResult", stdevResult);

            // serializationVerion 2
        }
    }

    [Serializable]
    public class SA_Report
    {
        public SA_CollObjectName report = new SA_CollObjectName();
        int serializeVersion = 1;

        public SA_Report(string coll, string reportName)
        {
            report.Coll = coll;
            report.objectName = reportName;
            report.saType.typeIdx = TypeIdx.eSAReport;
        }

        public string GetReportName()
        {
            return report.GetObjName();
        }

        public string GetQuickReportPDFFileName()
        {
            string ext = Path.GetExtension(GetReportName());
            if (ext.Length < ".pfd".Length || !GetReportName().Contains(".pdf"))
                ext = ".pdf";
            return report.GetObjName() + ext;
        }


        public SA_Report(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                report = new SA_CollObjectName(info, ctxt);
            }
        }

        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);

            report.GetObjectData(info, ctxt);
            // serializationVerion 2

        }
    }

    [Serializable]
    public class SA_InstID
    {
        public String Coll = @"A";
        public int instIdx = 0;
        public string Name = "";  // serialization = 2
        public string Model = ""; // serialization = 2

        int  serializeVersion = 2;

        public SA_InstID()
        {
        }

        public SA_InstID(String c, int idx)
        {
            Coll = c;
            instIdx = idx;
        }

        public String GetInstIDName()
        {
            String c_InstIDName = Coll + "::" + instIdx;
            return c_InstIDName;
        }

        public String GetColl()
        {
            return Coll;
        }

        public int GetInstIdx()
        {
            return instIdx;
        }

        public SA_InstID(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                Coll = info.GetString("Coll");
                instIdx = (int)info.GetValue("instIdx", typeof(int));
            }

            if (sVersion >= 2) // serializationVerion 2
            {
                Name = info.GetString("Name");
                Model = info.GetString("Model");
            }
        }

        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("Coll", Coll);
            info.AddValue("instIdx", instIdx);

            // serializationVerion 2
            info.AddValue("Name", Name);
            info.AddValue("Model", Model);
        }
    }

    [Serializable]
    public class SA_InstRefList
    {
        public List<SA_InstID> collInstRefList = new List<SA_InstID>();

        int serializeVersion = 1;


        public SA_InstRefList()
        {
            // typeName = typeNames.ElementAt(TypeIdx.eAny);
        }


        public SA_InstRefList(String name, int idx)
        {
            SA_InstID iid = new SA_InstID(name, idx);
            collInstRefList.Add(iid);
        }

        public String GetCollInstModel(int idx)
        {
            string model = "";
            if (collInstRefList.Count <= idx && idx >= 0)
            {
                model = collInstRefList[idx].Model;
            }
            return model;
        }

        public String GetCollInstNameIdx(int idx)
        {
            String c_InstName = "";
            if (collInstRefList.Count <= idx && idx >= 0)
            {
                c_InstName = collInstRefList[idx].Model;
            }
            return c_InstName;
        }

        public SA_InstID GetInstID(int idx)
        {
            SA_InstID inst = new SA_InstID();
            if (collInstRefList.Count <= idx && idx >= 0)
            {
                inst = collInstRefList[idx];
            }
            return inst;
        }

        public bool RemoveCollInstIdx(int idx)
        {
            bool ret = false;
            if (collInstRefList.Count <= idx && idx >= 0)
            {
                collInstRefList.RemoveAt(idx);
                ret = true;
            }
            return ret;
        }

        public void GetStringList(ref List<string> sl)
        {
            foreach (SA_InstID iid in collInstRefList)
            {
                sl.Add(iid.GetInstIDName());
            }
        }

        public SA_InstRefList(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                int count = (int)info.GetValue("count",typeof(int));
                for (int i = 0; i < count; i++)
                {
                    SA_InstID iid = new SA_InstID(info, ctxt);
                    collInstRefList.Add(iid);
                }
            }
        }

        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            info.AddValue("count", collInstRefList.Count);
            foreach (SA_InstID iid in collInstRefList)
            {
                iid.GetObjectData(info, ctxt);
            }

            // serializationVerion 2
        }
    }


    public class VectorGroupColorization
    {
        public String sColorRangeMethod = "Toleranced (Continuous)";
        public String sTopColor = "Blue";
        public String sMidColor = "Green";
        public String sBottomColor = "Red";
        public bool bTubes = true;
        public bool bDrawArrowheads = false;
        public bool bLableVectorsWithValues = true;
        public double fVectorMag = 500.000000;
        public int dLineWidth = 1;
        public bool bDrawBlotches = false;
        public double fBotchSize = 0.250000;
        public bool bShowOutOfTolOnly = true;
        public bool bShowColorBar = false;
        public bool bShowColorBarPercentages = false;
        public bool bShowColorBarFractions = false;
        public double fHiMag = 0.050000;
        public double fLoMag = -0.050000;
        public double fHiTol = 0.030000;
        public double fLoTol = -0.030000;

        public VectorGroupColorization(ToleranceSet ts)
        {
            fHiMag = ts.maxResult;
            fLoMag = -ts.maxResult;
            if (ts.maxResult < ts.rmsTol)
            {
                fHiMag = ts.rmsTol * 3;
                fLoMag = -ts.rmsTol * 3;
            }
            fHiTol = ts.rmsTol;
            fLoTol = -ts.rmsTol;

            if (ts.maxResult > 0.000001)
                fVectorMag = 1.0 / ts.maxResult;
        }
    }

    [Serializable]
    public class VectorGroupProperties
    {
        int serializeVersion = 1;
        public SA_CollObjectName VectorGroupName = new SA_CollObjectName(@"",@"VG",TypeIdx.eVectorGroup);
        
        public int totVectors = 0;
        public int numVectorsInTol = 0;
        public int numVectorsOutTol = 0;
        public double percentVectorInTol = 0.0;
        public double percentVectorOutTol = 0.0;
        public double absMaxMagResult = 0.0;
        public double absMinMagResult = 0.0;
        public double maxMag = 0.0;
        public double minMag = 0.0;
        public double stdev = 0.0;
        public double stdevMeanZero = 0.0;
        public double avgMag = 0.0;
        public double avgAbsMag = 0.0;
        public double highTolValue = 0.0;
        public double lowTolValue = 0.0;
        public double rmsValue = 0.0;

        public VectorGroupProperties()
        {
        }

        public VectorGroupProperties(SerializationInfo info, StreamingContext ctxt)
        {
            int sVersion = (int)info.GetValue("serializationVersion", typeof(int));

            // check serialization version

            if (sVersion >= 1) // serializationVerion 1
            {
                VectorGroupName = new SA_CollObjectName(info, ctxt);

                totVectors = (int)info.GetValue("totVectors",typeof(int));
                numVectorsInTol = (int)info.GetValue("newVectorsInTol", typeof(int));
                numVectorsOutTol = (int)info.GetValue("newVectorsOutTol", typeof(int));

                percentVectorInTol = info.GetDouble("percentVectorInTol");
                percentVectorOutTol = info.GetDouble("percentVectorOutTol");
                absMaxMagResult = info.GetDouble("absMaxMagResult");
                absMinMagResult = info.GetDouble("absMinMagResult");
                maxMag = info.GetDouble("maxMag");
                minMag = info.GetDouble("minMag");
                stdev = info.GetDouble("stdev");
                stdevMeanZero = info.GetDouble("stdevMeanZero");
                avgMag = info.GetDouble("avgMag");
                avgAbsMag = info.GetDouble("avgAbsMag");
                highTolValue = info.GetDouble("highTolValue");
                lowTolValue = info.GetDouble("lowTolValue");
                rmsValue = info.GetDouble("rmsValue");            
            }
        }

        //Serialization function.
        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            // serializationVerion 1
            info.AddValue("serializationVersion", serializeVersion);
            
            VectorGroupName.GetObjectData(info, ctxt);

            info.AddValue("totVectors", totVectors);
            info.AddValue("numVectorsInTol", numVectorsInTol);
            info.AddValue("numVectorsInTol", numVectorsInTol);

            info.AddValue("percentVectorInTol", percentVectorInTol);
            info.AddValue("percentVectorOutTol", percentVectorOutTol);
            info.AddValue("absMaxMagResult", absMaxMagResult);
            info.AddValue("absMinMagResult", absMinMagResult);

            info.AddValue("maxMag", maxMag);
            info.AddValue("minMag", minMag);
            info.AddValue("stdev", stdev);
            info.AddValue("stdevMeanZero", stdevMeanZero);

            info.AddValue("avgMag", avgMag);
            info.AddValue("avgAbsMag", avgAbsMag);

            info.AddValue("highTolValue", highTolValue);
            info.AddValue("lowTolValue", lowTolValue);
            info.AddValue("rmsValue", rmsValue);

            // serializationVerion 2
        }
    }


}
