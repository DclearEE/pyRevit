using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Text;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using XYZ = Autodesk.Revit.DB.XYZ;
using Microsoft.Office;
using Microsoft.Office.Interop.Excel;




namespace NBK.Spool
{
    public static class MicrodeskHelpers
    {
        public static string Log()
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\log.txt";
        }

        public static string GetVersionString()
        {
            //
            // Get the addin version
            //
            Version ver = Assembly.GetExecutingAssembly().GetName().Version;
            string[] parameters = { ver.Major.ToString(), ver.Minor.ToString(), ver.Build.ToString(), ver.Revision.ToString() };
            return String.Format(" v{0}.{1}.{2}.{3}", parameters);
        }

        public static void DisableTextRendering(System.Windows.Forms.Control control)
        {
            IEnumerable<System.Windows.Forms.Control> buttons = GetAllControls(control, typeof(System.Windows.Forms.Button));

            foreach (System.Windows.Forms.Control c in buttons)
            {
                ((System.Windows.Forms.Button)c).UseCompatibleTextRendering = false;
            }

            IEnumerable<System.Windows.Forms.Control> labels = GetAllControls(control, typeof(System.Windows.Forms.Label));

            foreach (System.Windows.Forms.Control c in labels)
            {
                ((System.Windows.Forms.Label)c).UseCompatibleTextRendering = false;
            }

            IEnumerable<System.Windows.Forms.Control> checkBoxes = GetAllControls(control, typeof(System.Windows.Forms.CheckBox));

            foreach (System.Windows.Forms.Control c in checkBoxes)
            {
                ((System.Windows.Forms.CheckBox)c).UseCompatibleTextRendering = false;
            }

            IEnumerable<System.Windows.Forms.Control> radioButtons = GetAllControls(control, typeof(System.Windows.Forms.RadioButton));

            foreach (System.Windows.Forms.Control c in radioButtons)
            {
                ((RadioButton)c).UseCompatibleTextRendering = false;
            }

            IEnumerable<System.Windows.Forms.Control> groupBoxes = GetAllControls(control, typeof(System.Windows.Forms.GroupBox));

            foreach (System.Windows.Forms.Control c in groupBoxes)
            {
                ((System.Windows.Forms.GroupBox)c).UseCompatibleTextRendering = false;
            }
        }

        public static IEnumerable<System.Windows.Forms.Control> GetAllControls(System.Windows.Forms.Control control, Type type)
        {
            var controls = control.Controls.Cast<System.Windows.Forms.Control>();

            return controls.SelectMany(ctrl => GetAllControls(ctrl, type))
                                      .Concat(controls)
                                      .Where(c => c.GetType() == type);
        }

        //
        //GENERAL TEXT
        //

        public static bool IsNumeric(string str)
        {
            bool textIsNum = true;
            try
            {
                int.Parse(str);
            }
            catch
            {
                textIsNum = false;
            }
            return textIsNum;
        }


        public static double GetSizeFromString(string input)
        {
            double diameter = 1;

            string[] inputArray = input.Split('[');
            string diamterString = inputArray[1];
            diamterString = diamterString.Replace(" ]", "");
            diamterString = diamterString.Replace(" ", "");

            switch (diamterString)
            {
                case "1/4":
                    diameter = .25;
                    break;
                case "1/2":
                    diameter = .5;
                    break;
                case "3/4":
                    diameter = .75;
                    break;
                case "1":
                    diameter = 1;
                    break;
                case "1-1/4":
                    diameter = 1.25;
                    break;
                case "1-1/2":
                    diameter = 1.5;
                    break;
                case "2":
                    diameter = 2;
                    break;
                case "2-1/2":
                    diameter = 2.5;
                    break;
                case "3":
                    diameter = 3;
                    break;
                case "4":
                    diameter = 4;
                    break;
                case "6":
                    diameter = 6;
                    break;
                case "8":
                    diameter = 8;
                    break;
                case "10":
                    diameter = 10;
                    break;
                case "12":
                    diameter = 12;
                    break;
                case "14":
                    diameter = 14;
                    break;
                case "16":
                    diameter = 16;
                    break;
                case "18":
                    diameter = 18;
                    break;
                case "20":
                    diameter = 20;
                    break;
                case "22":
                    diameter = 22;
                    break;
                case "24":
                    diameter = 24;
                    break;
                default:
                    diameter = 1;
                    break;
            }

            return diameter;
        }


        public static string[] ConvertToStringArray(System.Array values)
        {
            string[] newArray = new string[values.Length];

            int index = 0;
            for (int i = values.GetLowerBound(0);
                  i <= values.GetUpperBound(0); i++)
            {
                for (int j = values.GetLowerBound(1);
                          j <= values.GetUpperBound(1); j++)
                {
                    if (values.GetValue(i, j) == null)
                    {
                        newArray[index] = "";
                    }
                    else
                    {
                        newArray[index] = (string)values.GetValue(i, j).ToString();
                    }
                    index++;
                }
            }
            return newArray;
        }


        //
        //GENERAL MATH
        //

        /// <summary>
        /// Get the point on a line that would create a perpendicular branch to a point in space
        /// </summary>
        /// <param name="p0">XYZ start coordinate of line</param>
        /// <param name="p1">XYZ end coordinate of line</param>
        /// <param name="pX">XYZ point in space</param>
        /// <returns>GXY coordinate</returns>
        public static XYZ PerpIntersection(XYZ p0, XYZ p1, XYZ pX)
        {
            //We're going to use formulas to find the minimum distance between a point and  a line.
            //Calculate the dot product of the vectors
            double u = ((pX.X - p0.X) * (p1.X - p0.X)) + ((pX.Y - p0.Y) * (p1.Y - p0.Y)) + ((pX.Z - p0.Z) * (p1.Z - p0.Z));
            double dist = p0.DistanceTo(p1);
            u = u / (dist * dist);

            XYZ t = new XYZ(p0.X + u * (p1.X - p0.X), p0.Y + u * (p1.Y - p0.Y), p0.Z + u * (p1.Z - p0.Z));

            return t;
        }


        public static XYZ Midpoint(XYZ p0, XYZ p1)
        {
            XYZ t = new XYZ((p0.X + p1.X) / 2, (p0.Y + p1.Y) / 2, (p0.Z + p1.Z) / 2);

            return t;
        }


        public static XYZ Centroid(XYZ p0, XYZ p1, XYZ p2)
        {
            XYZ t = new XYZ((p0.X + p1.X + p2.X) / 3, (p0.Y + p1.Y + p2.Y) / 3, (p0.Z + p1.Z + p2.Z) / 3);

            return t;
        }


        /// <summary>
        /// Get the coordinate intersection of two vectors
        /// </summary>
        /// <param name="upstreamBranch">XYZ start coordinate of first line</param>
        /// <param name="downstreamBranch">XYZ end coordinate of first line</param>
        /// <param name="downstreamMain">XYZ start coordinate of second line</param>
        /// <param name="upstreamMain">XYZ end coordinate of second line</param>
        /// <returns></returns>
        public static XYZ IntersectionTwoVectors(XYZ upstreamBranch, XYZ downstreamBranch, XYZ downstreamMain, XYZ upstreamMain)
        {
            //Calculate the point where the two vectors are closest to each other
            //http://geomalgorithms.com/a07-_distance.html

            //Get a point on each vector
            XYZ p0 = downstreamMain;
            XYZ q0 = downstreamBranch;

            //Get the direction of each vector
            XYZ u = new XYZ(upstreamMain.X - downstreamMain.X,
                                upstreamMain.Y - downstreamMain.Y,
                                upstreamMain.Z - downstreamMain.Z);
            XYZ v = new XYZ(upstreamBranch.X - downstreamBranch.X,
                                upstreamBranch.Y - downstreamBranch.Y,
                                upstreamBranch.Z - downstreamBranch.Z);
            XYZ w = new XYZ(downstreamMain.X - downstreamBranch.X,
                                downstreamMain.Y - downstreamBranch.Y,
                                downstreamMain.Z - downstreamBranch.Z);

            double a = u.DotProduct(u);
            double b = u.DotProduct(v);
            double c = v.DotProduct(v);
            double d = u.DotProduct(w);
            double e = v.DotProduct(w);
            double D = a * c - b * b;
            double sc;
            double tc;

            //If the lines are almost parallel
            if (D < 0.00000001)
            {
                sc = 0;
                if (b > c)
                {
                    tc = d / b;
                }
                else
                {
                    tc = e / c;
                }
            }
            else
            {
                sc = (b * e - c * d) / D;
                tc = (a * e - b * d) / D;
            }

            //Solve for point P(Sc) = P0 + Sc*u
            XYZ intersect1 = new XYZ(p0.X + sc * u.X, p0.Y + sc * u.Y, p0.Z + sc * u.Z);
            //Solve for point Q(Tc) = Q0 + Tc*v
            XYZ intersect2 = new XYZ(q0.X + tc * v.X, q0.Y + tc * v.Y, q0.Z + tc * v.Z);

            XYZ finalIntersect = new XYZ((intersect2.X - intersect1.X) / 2 + intersect1.X, (intersect2.Y - intersect1.Y) / 2 + intersect1.Y, (intersect2.Z - intersect1.Z) / 2 + intersect1.Z);

            return finalIntersect;
        }




        /// <summary>
        /// Get a point projected from a 3D coordinate onto a 2D plane.
        /// </summary>
        /// <param name="planeNormal">XYZ Unit Vector perpendicular to the plane</param>
        /// <param name="pointOnPlane">XYZ coordinate anywhere on the plane</param>
        /// <param name="pointToProject">XYZ coordinate</param>
        /// <returns>XYZ projected point</returns>
        public static XYZ ProjectPointOnPlane(XYZ planeNormal, XYZ pointOnPlane, XYZ pointToProject)
        {

            double distance;
            XYZ translationVector;

            //First calculate the distance from the point to the plane:
            distance = SignedDistancePlanePoint(planeNormal, pointOnPlane, pointToProject);

            //Reverse the sign of the distance
            distance *= -1;

            //Get a translation vector
            translationVector = SetVectorLength(planeNormal, distance);

            //Translate the point to form a projection
            return pointToProject + translationVector;
        }


        /// <summary>
        /// Get the shortest distance between a point and a plane.
        /// The output is signed so it holds information as to which side of the plane normal the point is.
        /// </summary>
        /// <param name="planeNormal">XYZ Unit Vector perpendicular to the plane</param>
        /// <param name="pointOnPlane">XYZ coordinate anywhere on the plane</param>
        /// <param name="point">XYZ coordinate</param>
        /// <returns>XYZ nearest coordinate</returns>
        public static double SignedDistancePlanePoint(XYZ planeNormal, XYZ pointOnPlane, XYZ point)
        {
            //dot product of the plane's normal vector and the difference between the point and the origin of the plane
            return planeNormal.X * (point.X - pointOnPlane.X) + planeNormal.Y * (point.Y - pointOnPlane.Y) + planeNormal.Z * (point.Z - pointOnPlane.Z);
        }


        //create a vector of direction "vector" with length "size"
        public static XYZ SetVectorLength(XYZ vector, double size)
        {
            //normalize the vector
            XYZ vectorNormalized = vector.Normalize();

            //scale the vector
            return vectorNormalized *= size;
        }



        /// <summary>
        /// Create a vector normal to a plane containing 3 points
        /// </summary>
        /// <param name="p">XYZ coordinate anywhere on the plane</param>
        /// <param name="q">XYZ coordinate anywhere on the plane</param>
        /// <param name="r">XYZ coordinate anywhere on the plane</param>
        /// <returns></returns>
        public static XYZ NormalThreePoints(XYZ p, XYZ q, XYZ r)
        {
            XYZ pq = q - p;
            XYZ pr = r - p;

            return pq.CrossProduct(pr);
        }

        /// <summary>
        /// Get the intersection of a line on a plane
        /// </summary>
        /// <param name="planeNormal">XYZ Unit Vector perpendicular to the plane</param>
        /// <param name="pointOnPlane">XYZ coordinate anywhere on the plane</param>
        /// <param name="p0">XYZ start coordinate of the line</param>
        /// <param name="p1">XYZ end coordinate of the line</param>
        /// <returns></returns>
        public static XYZ GetIntersectionOnPlane(XYZ planeNormal, XYZ pointOnPlane, XYZ p0, XYZ p1)
        {
            XYZ u = (p1 - p0);
            XYZ w = p0 - pointOnPlane;

            double d = planeNormal.DotProduct(u);
            double n = -planeNormal.DotProduct(w);

            if (Math.Abs(Math.Round(d, 3)) == 1)
            {
                return null;
            }

            double sI = n / d;

            if (sI < 0 || sI > 1)
            {
                return null;
            }


            return p0 + sI * u;
        }



        //
        //REVIT
        //




        public static Element GetFamilyType(string name, BuiltInCategory bic, Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(bic);
            collector.OfClass(typeof(ElementType));

            //Identify Pipe Family Type
            foreach (Element etype in collector)
            {
                if (etype.Name == name)
                {
                    return etype;
                }
            }
            return null;
        }


        public static Element GetLevelByName(string name, Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Levels);
            collector.OfClass(typeof(Level));

            //Identify Pipe Family Type
            foreach (Element e in collector)
            {
                if (e.Name == name)
                {
                    return e;
                }
            }
            return null;
        }


        public static ElementId GetLevel(Document doc, double z)
        {
            Level level = null;
            ElementId LevelId = null;


            //
            // Try first to get the level of the active view
            //
            try
            {
                LevelId = doc.ActiveView.GenLevel.Id;
            }
            //
            // If the active view has no associated level, in the case of a 3D view, get the closest level below the z coordinate.
            //
            catch
            {
                //
                // Collect all the Levels in the document
                //
                FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(Level));

                if(collector.Count() == 0)
                    TaskDialog.Show("Warning", "The active model contains no Levels.  Levels are required to create floor plans.");
                

                //
                // Cast the Element collection to Levels and order them by their elevations
                //
                List<Level> levels = collector.Cast<Level>().ToList<Level>();
                List<Level> orderedLevels = levels.OrderBy(w => w.Elevation).ToList();

                //
                // Check if the given z coordinate is below the lowest level. 
                // If so, the pipe is underground and the first level is correct.
                //
                if (z < orderedLevels.First().Elevation) return orderedLevels.First().Id;

                //
                // Find the first level below the z coordinate
                //
                return orderedLevels.Where(w => w.Elevation <= z).Last().Id;
            }

            return LevelId;
        }


        public static PipingSystemType GetPipeSystem(Document doc, Connector connector)
        {
            //
            //Skip non-pipe connectors
            //
            if (connector.Domain != Domain.DomainPiping)
            {
                return null;
            }

            //
            //If the connector already has a sytem applied, get its PipingSystemType
            //
            try
            {
                PipingSystem pSys = connector.MEPSystem as PipingSystem;
                ElementId sysId = pSys.GetTypeId();
                PipingSystemType sys = doc.GetElement(sysId) as PipingSystemType;
                return sys;
            }
            catch { }

            //
            //Otherwise, collect all the PipingSystemTypes in the model and 
            //
            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(PipingSystemType));

            string pipeSystemTypeString = connector.PipeSystemType.ToString();

            if (pipeSystemTypeString.Equals("UndefinedSystemType") ||
               pipeSystemTypeString.Equals("Fitting") ||
               pipeSystemTypeString.Equals("Global"))
            {
                //
                //set the first one as a default
                //
                return collector.First() as PipingSystemType;
            }
            else
            {
                string pipeSysClassString = "Hydronic Supply";
                switch (pipeSystemTypeString)
                {
                    case "SupplyHydronic":
                        pipeSysClassString = "Hydronic Supply";
                        break;
                    case "ReturnHydronic":
                        pipeSysClassString = "Hydronic Return";
                        break;
                    case "Sanitary":
                        pipeSysClassString = "Sanitary";
                        break;
                    case "Vent":
                        pipeSysClassString = "Vent";
                        break;
                    case "DomesticHotWater":
                        pipeSysClassString = "Domestic Hot Water";
                        break;
                    case "DomesticColdWater":
                        pipeSysClassString = "Domestic Cold Water";
                        break;
                    case "OtherPipe":
                        pipeSysClassString = "Other";
                        break;
                    case "FireProtectWet":
                        pipeSysClassString = "Fire Protection Wet";
                        break;
                    case "FireProtectDry":
                        pipeSysClassString = "Fire Protection Dry";
                        break;
                    case "FireProtectPreaction":
                        pipeSysClassString = "Fire Protection Preaction";
                        break;
                    case "FireProtectOther":
                        pipeSysClassString = "Fire Protection Other";
                        break;
                    default:
                        pipeSysClassString = "Hydronic Supply";
                        break;
                }

                foreach (PipingSystemType sys in collector)
                {
                    if (sys.LookupParameter("System Classification").AsString().Equals(pipeSysClassString))
                    {
                        return sys;
                    }
                }
            }

            return collector.First() as PipingSystemType;
        }


        public static PipeInsulationType GetPipeInsulation(Document doc)
        {
            FilteredElementCollector fec = new FilteredElementCollector(doc).OfClass(typeof(PipeInsulationType));

            PipeInsulationType pipeInsulation = null;

            foreach (PipeInsulationType pi in fec)
            {
                pipeInsulation = pi;
            }
            return pipeInsulation;
        }


        public static Connector[] ConnectorArray(Element e)
        {
            if (e == null) return null;
            FamilyInstance fi = e as FamilyInstance;
            ConnectorSet cS = null;
            List<Connector> cList = new List<Connector>();

            if (fi != null && fi.MEPModel != null)
            {
                cS = fi.MEPModel.ConnectorManager.Connectors;
            }

            MEPSystem system = e as MEPSystem;
            if (system != null)
            {
                cS = system.ConnectorManager.Connectors;
            }

            MEPCurve ductorPipe = e as MEPCurve;
            if (ductorPipe != null)
            {
                cS = ductorPipe.ConnectorManager.Connectors;
            }

            FabricationPart fabPart = e as FabricationPart;
            if (fabPart != null)
            {
                cS = fabPart.ConnectorManager.Connectors;
            }

            if (cS == null) return null;
            foreach (Connector c in cS)
            {
                cList.Add(c);
            }

            return cList.ToArray();

        }


        public static Connector[] UnusedConnectorArray(Element e)
        {
            if (e == null) return null;
            FamilyInstance fi = e as FamilyInstance;
            ConnectorSet cS = null;
            List<Connector> cList = new List<Connector>();

            if (fi != null && fi.MEPModel != null)
            {
                cS = fi.MEPModel.ConnectorManager.UnusedConnectors;
            }

            MEPSystem system = e as MEPSystem;
            if (system != null)
            {
                cS = system.ConnectorManager.UnusedConnectors;
            }

            MEPCurve ductorPipe = e as MEPCurve;
            if (ductorPipe != null)
            {
                cS = ductorPipe.ConnectorManager.UnusedConnectors;
            }

            if (cS == null) return null;
            foreach (Connector c in cS)
            {
                cList.Add(c);
            }

            return cList.ToArray();

        }


        public static Connector ClosestConnector(Connector[] cA1, Connector[] cA2)
        {
            Connector closestConnector = null;
            double d = double.MaxValue;
            double shortestLineLength = double.MaxValue;
            //Line testLine;
            int m = 0;
            foreach (Connector c1 in cA1)
            {

                if (cA1[m].Origin != cA2[0].Origin)
                {
                    d = Math.Sqrt(Math.Pow((cA1[m].Origin.X - cA2[0].Origin.X), 2) + Math.Pow((cA1[m].Origin.Y - cA2[0].Origin.Y), 2) + Math.Pow((cA1[m].Origin.Z - cA2[0].Origin.Z), 2));
                    if (d < shortestLineLength)
                    {
                        shortestLineLength = d;
                        closestConnector = cA1[m];
                    }
                }
                m++;
            }

            return closestConnector;
        }


        public static Connector ClosestAvailableConnector(Connector[] cA1, Connector[] cA2)
        {
            Connector closestConnector = null;
            double d = double.MaxValue;
            double shortestLineLength = double.MaxValue;
            //Line testLine;
            int m = 0;
            foreach (Connector c1 in cA1)
            {
                //
                //Ignore connected connectors
                //
                if (c1.IsConnected)
                {
                    m++;
                    continue;
                }
                if (cA1[m].Origin != cA2[0].Origin)
                {
                    d = Math.Sqrt(Math.Pow((cA1[m].Origin.X - cA2[0].Origin.X), 2) + Math.Pow((cA1[m].Origin.Y - cA2[0].Origin.Y), 2) + Math.Pow((cA1[m].Origin.Z - cA2[0].Origin.Z), 2));
                    if (d < shortestLineLength)
                    {
                        shortestLineLength = d;
                        closestConnector = cA1[m];
                    }
                }
                else
                {
                    //
                    //In this case we found 2 connectors at exaclty the same point
                    //
                    return c1;
                }
                m++;
            }

            return closestConnector;
        }


        public static Connector ClosestConnectorOfDomain(Connector[] cA1, Connector[] cA2, Domain domain)
        {
            Connector closestConnector = null;
            double d = double.MaxValue;
            double shortestLineLength = double.MaxValue;
            //Line testLine;
            int m = 0;
            foreach (Connector c1 in cA1)
            {
                if (c1.IsConnected)
                {
                    m++;
                    continue;
                }

                if (cA1[m].Origin != cA2[0].Origin)
                {
                    d = Math.Sqrt(Math.Pow((cA1[m].Origin.X - cA2[0].Origin.X), 2) + Math.Pow((cA1[m].Origin.Y - cA2[0].Origin.Y), 2) + Math.Pow((cA1[m].Origin.Z - cA2[0].Origin.Z), 2));
                    if (d < shortestLineLength && cA1[m].Domain == domain)
                    {
                        shortestLineLength = d;
                        closestConnector = cA1[m];
                    }
                }
                m++;
            }

            return closestConnector;
        }


        public static Connector ClosestConnectorOfDomainAndAngle(Connector[] cA1, Connector c)
        {
            Connector closestConnector = null;
            double d = double.MaxValue;
            double shortestLineLength = double.MaxValue;
            //Line testLine;
            int m = 0;
            foreach (Connector c1 in cA1)
            {
                if (c1.IsConnected)
                {
                    m++;
                    continue;
                }

                if (cA1[m].Origin != c.Origin)
                {
                    d = Math.Sqrt(Math.Pow((cA1[m].Origin.X - c.Origin.X), 2) + Math.Pow((cA1[m].Origin.Y - c.Origin.Y), 2) + Math.Pow((cA1[m].Origin.Z - c.Origin.Z), 2));

                    if (d < shortestLineLength //Get the closest connector
                        && cA1[m].Domain == c.Domain //Make sure it belongs to the same domain (i.e. duct, pipe)
                        && c.CoordinateSystem.BasisZ.DotProduct(cA1[m].CoordinateSystem.BasisZ) < -.9 //Make sure the connectors are not pointing in the opposite direction
                        )
                    {
                        shortestLineLength = d;
                        closestConnector = cA1[m];
                    }
                }
                m++;
            }

            return closestConnector;
        }


        public static Connector NearestConnector(Connector[] cA, XYZ startPoint)
        {
            Connector closestConnector = null;
            try
            {
                double d = double.MaxValue;
                double shortestLineLength = double.MaxValue;
                //Line testLine;
                int m = 0;
                foreach (Connector c1 in cA)
                {
                    d = Math.Sqrt(Math.Pow((cA[m].Origin.X - startPoint.X), 2) + Math.Pow((cA[m].Origin.Y - startPoint.Y), 2) + Math.Pow((cA[m].Origin.Z - startPoint.Z), 2));
                    if (d < shortestLineLength)
                    {

                        shortestLineLength = d;
                        closestConnector = cA[m];
                    }
                    m++;
                }
            }
            catch { }
            return closestConnector;
        }


        public static Connector[] NearestConnectors(List<Connector> cs, XYZ basePoint)
        {
            List<Connector> closestConnectors = new List<Connector>();

            foreach (Connector c1 in cs)
            {
                if (Math.Round(c1.Origin.X, 2) == Math.Round(basePoint.X, 2)
                    && Math.Round(c1.Origin.Y, 2) == Math.Round(basePoint.Y, 2))
                {
                    closestConnectors.Add(c1);
                }
            }

            return closestConnectors.ToArray();
        }


        public static Connector FarthestConnector(Connector[] cA, XYZ startPoint)
        {
            Connector farthestConnector = null;
            try
            {
                double d = double.MinValue;
                double farthestLineLength = double.MinValue;
                //Line testLine;
                int m = 0;
                foreach (Connector c1 in cA)
                {
                    d = Math.Sqrt(Math.Pow((cA[m].Origin.X - startPoint.X), 2) + Math.Pow((cA[m].Origin.Y - startPoint.Y), 2) + Math.Pow((cA[m].Origin.Z - startPoint.Z), 2));
                    if (d > farthestLineLength)
                    {

                        farthestLineLength = d;
                        farthestConnector = cA[m];
                    }
                    m++;
                }
            }
            catch { }
            return farthestConnector;
        }


        public static Connector ConnectedConnector(Connector cInput, XYZ origin)
        {
            Connector attachedConnector = null;
            ConnectorSet attachedConnectors = cInput.AllRefs;

            foreach (Connector c in attachedConnectors)
            {
                //The ConnectorSet will return a logical connection as well as one physical one
                if (c.ConnectorType != ConnectorType.End)
                {
                    //This is the logical connector, so skip it
                    continue;
                }


                //Otherwise, if we find the physical connector, make sure it's connected.
                //If not, then abort.
                if (c.IsConnected == false)
                {
                    continue;
                }

                if (Math.Round(c.Origin.X, 2) == Math.Round(origin.X, 2) &&
                    Math.Round(c.Origin.Y, 2) == Math.Round(origin.Y, 2))
                {
                    //Once we find a physical connector that is connected, we can carry on...
                    attachedConnector = c;
                }
            }

            return attachedConnector;
        }


        public static FamilySymbol GetFittingFromRouting(PipeType pipeType, RoutingPreferenceRuleGroupType groupType, double pipeDiameter, string log, Document doc)
        {
            FamilySymbol fittingFamily = null;

            //
            //Find the routing preferences of the pipe
            //

            RoutingPreferenceManager rPM = pipeType.RoutingPreferenceManager;

            ElementId fittingTypeID = null;

            #region Log
            using (StreamWriter writer = new StreamWriter(log, true))
            {
                writer.WriteLine("rPM.GetNumberOfRules(group) {0}", rPM.GetNumberOfRules(groupType));
            }
            #endregion Log

            for (int k = 0; k != rPM.GetNumberOfRules(groupType); ++k)
            {
                #region Log
                using (StreamWriter writer = new StreamWriter(log, true))
                {
                    writer.WriteLine("rule {0}", k);
                }
                #endregion Log

                RoutingPreferenceRule rule = rPM.GetRule(groupType, k);

                PrimarySizeCriterion psc = rule.GetCriterion(0) as PrimarySizeCriterion;

                #region Log
                using (StreamWriter writer = new StreamWriter(log, true))
                {
                    writer.WriteLine("width: {0}", pipeDiameter);
                    writer.WriteLine("psc.MinimumSize: {0}", psc.MinimumSize * 12);
                    writer.WriteLine("psc.MaximumSize: {0}", psc.MaximumSize * 12);
                }
                #endregion Log

                if (Math.Round((psc.MinimumSize * 12), 2) <= Math.Round(pipeDiameter, 2) &&
                    Math.Round(pipeDiameter, 2) <= Math.Round((psc.MaximumSize * 12), 2))
                {
                    fittingTypeID = rule.MEPPartId;

                    fittingFamily = doc.GetElement(fittingTypeID) as FamilySymbol;

                    #region Log
                    using (StreamWriter writer = new StreamWriter(log, true))
                    {
                        writer.WriteLine("fittingFamily.Name: {0}", fittingFamily.Name);
                    }
                    #endregion Log

                    if (!fittingFamily.Name.Contains("Tap"))
                    {
                        //In this case we found something that meets the size requirement and is not a tap
                        break;
                    }
                }
            }
            return fittingFamily;
        }


        public static FamilySymbol GetFittingFromConduit(ConduitType condType, RoutingPreferenceRuleGroupType groupType, double pipeDiameter, string log, Document doc)
        {
            FamilySymbol fittingFamily = null;

            //
            //Find the routing preferences of the pipe
            //

            RoutingPreferenceManager rPM = condType.RoutingPreferenceManager;

            ElementId fittingTypeID = null;

            #region Log
            using (StreamWriter writer = new StreamWriter(log, true))
            {
                writer.WriteLine("rPM.GetNumberOfRules(group) {0}", rPM.GetNumberOfRules(groupType));
            }
            #endregion Log

            for (int k = 0; k != rPM.GetNumberOfRules(groupType); ++k)
            {
                #region Log
                using (StreamWriter writer = new StreamWriter(log, true))
                {
                    writer.WriteLine("rule {0}", k);
                }
                #endregion Log

                RoutingPreferenceRule rule = rPM.GetRule(groupType, k);

                PrimarySizeCriterion psc = rule.GetCriterion(0) as PrimarySizeCriterion;

                #region Log
                using (StreamWriter writer = new StreamWriter(log, true))
                {
                    writer.WriteLine("width: {0}", pipeDiameter);
                    writer.WriteLine("psc.MinimumSize: {0}", psc.MinimumSize * 12);
                    writer.WriteLine("psc.MaximumSize: {0}", psc.MaximumSize * 12);
                }
                #endregion Log

                if (Math.Round((psc.MinimumSize * 12), 2) <= Math.Round(pipeDiameter, 2) &&
                    Math.Round(pipeDiameter, 2) <= Math.Round((psc.MaximumSize * 12), 2))
                {
                    fittingTypeID = rule.MEPPartId;

                    fittingFamily = doc.GetElement(fittingTypeID) as FamilySymbol;

                    #region Log
                    using (StreamWriter writer = new StreamWriter(log, true))
                    {
                        writer.WriteLine("fittingFamily.Name: {0}", fittingFamily.Name);
                    }
                    #endregion Log

                    if (!fittingFamily.Name.Contains("Tap"))
                    {
                        //In this case we found something that meets the size requirement and is not a tap
                        break;
                    }
                }
            }
            return fittingFamily;
        }


        public static double RotateXY(XYZ insertion, FamilyInstance newFitting, double deltaX, double deltaY, string log, Document doc)
        {
            double angleXY;
            //Pipe draining North
            if (Math.Round(deltaX, 4) == 0 &&
                Math.Round(deltaY, 4) < 0)
            {
                #region Log
                using (StreamWriter writer = new StreamWriter(log, true))
                {
                    writer.WriteLine("Pipe draining North");
                }
                #endregion Log
                angleXY = 3 * Math.PI / 2;
            }
            //Pipe draining South
            else if (Math.Round(deltaX, 4) == 0 &&
                    Math.Round(deltaY, 4) > 0)
            {
                #region Log
                using (StreamWriter writer = new StreamWriter(log, true))
                {
                    writer.WriteLine("Pipe draining South");
                }
                #endregion Log
                angleXY = Math.PI / 2;
            }
            //Pipe draining East
            else if (Math.Round(deltaX, 4) < 0 &&
                Math.Round(deltaY, 4) == 0)
            {
                #region Log
                using (StreamWriter writer = new StreamWriter(log, true))
                {
                    writer.WriteLine("Pipe draining East");
                }
                #endregion Log
                angleXY = Math.PI;
            }
            //Pipe draining West
            else if (Math.Round(deltaX, 4) > 0 &&
                Math.Round(deltaY, 4) == 0)
            {
                #region Log
                using (StreamWriter writer = new StreamWriter(log, true))
                {
                    writer.WriteLine("Pipe draining West");
                }
                #endregion Log
                angleXY = 0;
            }
            else if (Math.Round(deltaX, 4) < 0)
            {
                #region Log
                using (StreamWriter writer = new StreamWriter(log, true))
                {
                    writer.WriteLine("deltaX < 0");
                }
                #endregion Log
                angleXY = Math.PI + Math.Atan(deltaY / deltaX);
            }
            else
            {
                #region Log
                using (StreamWriter writer = new StreamWriter(log, true))
                {
                    writer.WriteLine("last chance");
                }
                #endregion Log
                angleXY = Math.Atan(deltaY / deltaX);
            }

            //
            //rotate the fitting on the X-Y plane
            //
            XYZ perpEnd2 = new XYZ(insertion.X, insertion.Y, insertion.Z + 1);
            Autodesk.Revit.DB.Line axis2 = Autodesk.Revit.DB.Line.CreateBound(insertion, perpEnd2);
            ElementTransformUtils.RotateElement(doc, newFitting.Id, axis2, angleXY);
            return angleXY;
        }


        public static XYZ[] CalcBoundingBox(Autodesk.Revit.ApplicationServices.Application app, Document doc, Connector[] connectors, double height, double width)
        {
            XYZ[] bbCorners = new XYZ[2];

            bbCorners[0] = connectors[0].Origin + connectors[0].CoordinateSystem.BasisX * width / 2.0 + connectors[0].CoordinateSystem.BasisY * height / 2.0;
            bbCorners[1] = connectors[1].Origin + connectors[1].CoordinateSystem.BasisX * width / 2.0 - connectors[1].CoordinateSystem.BasisY * height / 2.0;

            // Make sure we have a proper min max
            double minX = bbCorners[0].X;
            double maxX = bbCorners[1].X;
            double minY = bbCorners[0].Y;
            double maxY = bbCorners[1].Y;
            double minZ = bbCorners[0].Z;
            double maxZ = bbCorners[1].Z;

            if (bbCorners[0].X > bbCorners[1].X)
            {
                minX = bbCorners[1].X;
                maxX = bbCorners[0].X;
            }
            if (bbCorners[0].Y > bbCorners[1].Y)
            {
                minY = bbCorners[1].Y;
                maxY = bbCorners[0].Y;
            }
            if (bbCorners[0].Z > bbCorners[1].Z)
            {
                minZ = bbCorners[1].Z;
                maxZ = bbCorners[0].Z;
            }
            bbCorners[0] = new XYZ(minX, minY, minZ);
            bbCorners[1] = new XYZ(maxX, maxY, maxZ);

            return bbCorners;
        }

        public static void RotateOnProjectedPlane(XYZ planeNormEndPoint1, XYZ planeNormEndPoint2, XYZ correctCoord, FamilyInstance newFitting, Connector[] connArray, double additionalAng, string log, Document doc, SubTransaction sT)
        {
            //
            //We want to rotate the fitting along the centerline of its first connector
            //That centerline can be defined by the first coordinate and the insertion point of the elbow.
            //To do that, we need to project the points of the fitting's branch and the nearest connector of the branch onto a 2D plane that is perpendicular to the centerline of the wye
            //
            //Get the new positions of the elbow endpoints

            Autodesk.Revit.DB.Line centerLine = Autodesk.Revit.DB.Line.CreateBound(planeNormEndPoint1, planeNormEndPoint2);


            //Get a unit vector of the difference between the endpoints.  This will be a vector orthagonal to the plane we need.
            XYZ planeNormal = (planeNormEndPoint2 - planeNormEndPoint1).Normalize();

            #region Log
            using (StreamWriter writer = new StreamWriter(log, true))
            {
                writer.WriteLine("planeNormal: {0}, {1}, {2}", Math.Round(planeNormal.X, 2), Math.Round(planeNormal.Y, 2), Math.Round(planeNormal.Z, 2));
            }
            #endregion Log

            //project the ends of the elbow on a 2D plane normal to the elbow
            //To determine which end of the elbow we want to use as a base for the rotation angle,
            //let's project both ends on the 2D plane and determine which is further from the insertion point
            XYZ elbowEnd2D = MicrodeskHelpers.ProjectPointOnPlane(planeNormal, planeNormEndPoint1, connArray[0].Origin);
            XYZ elbowEnd2Da = MicrodeskHelpers.ProjectPointOnPlane(planeNormal, planeNormEndPoint1, connArray[1].Origin);

            #region Log
            using (StreamWriter writer = new StreamWriter(log, true))
            {
                writer.WriteLine("elbowEnd2Da: {0}", Math.Abs((elbowEnd2Da - planeNormEndPoint1).GetLength()));
                writer.WriteLine("elbowEnd2D: {0}", Math.Abs((elbowEnd2D - planeNormEndPoint1).GetLength()));
            }
            #endregion Log

            Connector conn = connArray[0];

            if ((elbowEnd2Da - planeNormEndPoint1).GetLength() > (elbowEnd2D - planeNormEndPoint1).GetLength())
            {
                elbowEnd2D = elbowEnd2Da;
                conn = connArray[1];
            }

            XYZ correctCoord2D = MicrodeskHelpers.ProjectPointOnPlane(planeNormal, planeNormEndPoint1, correctCoord);


            //Get the magnitude of the vector between the correctd coordinate (projected on the 2D plane) and the elbow endpoint (on the 2D plane)
            double displacement = Math.Abs((elbowEnd2D - correctCoord2D).GetLength());


            //Convert the two points to vectors using the insertion point as their intersection
            XYZ elbowVector = new XYZ(elbowEnd2D.X - planeNormEndPoint1.X, elbowEnd2D.Y - planeNormEndPoint1.Y, elbowEnd2D.Z - planeNormEndPoint1.Z);
            XYZ correctCoordVector = new XYZ(correctCoord2D.X - planeNormEndPoint1.X, correctCoord2D.Y - planeNormEndPoint1.Y, correctCoord2D.Z - planeNormEndPoint1.Z);



            //Calculate the magnitude of each vector
            double elbowVectorM = elbowVector.GetLength();
            double coordVectorM = correctCoordVector.GetLength();




            double planeDirection = 1;
            if (planeNormal.X + planeNormal.Y + planeNormal.Z < 0)
            {
                planeDirection = -1;
            }

            //The angle = (elbowVector dot branchVector) / (elbowVector * branchVectorM)
            //double ang = -Math.Acos(wyeVector.DotProduct(branchVector) / (wyeVectorM * branchVectorM));
            double ang = Math.Acos(elbowVector.DotProduct(correctCoordVector) / (elbowVectorM * coordVectorM)) * planeDirection;

            sT.Start();
            ElementTransformUtils.RotateElement(doc, newFitting.Id, centerLine, ang);
            sT.Commit();

            //it is possible we rotated the fitting in the opposite direction.
            //To check, verify the magnitude of the new elbow vector X correct coordinate vector is 0

            XYZ elbowEnd2D2 = MicrodeskHelpers.ProjectPointOnPlane(planeNormal, planeNormEndPoint1, conn.Origin);
            XYZ newElbowVector = new XYZ(elbowEnd2D2.X - planeNormEndPoint1.X, elbowEnd2D2.Y - planeNormEndPoint1.Y, elbowEnd2D2.Z - planeNormEndPoint1.Z);
            double displacement2 = Math.Abs((elbowEnd2D2 - correctCoord2D).GetLength());

            XYZ crossProd = newElbowVector.CrossProduct(correctCoordVector);
            double crossProdCheck = crossProd.GetLength();

            #region Log
            using (StreamWriter writer = new StreamWriter(log, true))
            {
                writer.WriteLine("elbowEnd2D: {0}, {1}, {2}", Math.Round(elbowEnd2D.X, 2), Math.Round(elbowEnd2D.Y, 2), Math.Round(elbowEnd2D.Z, 2));
                writer.WriteLine("correctCoord2D: {0}, {1}, {2}", Math.Round(correctCoord2D.X, 2), Math.Round(correctCoord2D.Y, 2), Math.Round(correctCoord2D.Z, 2));
                writer.WriteLine("elbowVector: {0}, {1}, {2}", Math.Round(elbowVector.X, 2), Math.Round(elbowVector.Y, 2), Math.Round(elbowVector.Z, 2));
                writer.WriteLine("correctCoordVector: {0}, {1}, {2}", Math.Round(correctCoordVector.X, 2), Math.Round(correctCoordVector.Y, 2), Math.Round(correctCoordVector.Z, 2));
                writer.WriteLine("elbowVectorM: {0}", elbowVectorM);
                writer.WriteLine("coordVectorM: {0}", coordVectorM);
                writer.WriteLine("Rotation Angle about Centerline (rads): {0}", ang);
                writer.WriteLine("Rotation Angle about Centerline (degs): {0}", ang * 180 / Math.PI);
                writer.WriteLine("crossProd: {0}, {1}, {2}", Math.Round(crossProd.X, 2), Math.Round(crossProd.Y, 2), Math.Round(crossProd.Z, 2));
                writer.WriteLine("crossProdCheck: {0}", Math.Round(crossProdCheck, 6));
                writer.WriteLine("displacement: {0}", displacement);
                writer.WriteLine("displacement2: {0}", displacement2);
            }
            #endregion Log


            if (Math.Round(crossProdCheck, 6) != 0 || displacement2 > displacement)
            {
                //We rotated the fitting in the opposite direction
                ang = -2 * ang;

                #region Log
                using (StreamWriter writer = new StreamWriter(log, true))
                {
                    writer.WriteLine("The new location of the connector is further away from the correct coordinate than when we started");
                    writer.WriteLine("New Angle about Centerline (rads): {0}", ang);
                    writer.WriteLine("New Angle about Centerline (degs): {0}", ang * 180 / Math.PI);
                }
                #endregion Log

                sT.Start();
                ElementTransformUtils.RotateElement(doc, newFitting.Id, centerLine, ang);
                sT.Commit();
            }
        }


        public static List<Element> FindMateNearby(List<Element> elementsFiltered, BuiltInCategory bic, XYZ point, int noIterations, string log, Document doc)
        {
            for (int j = 0; j <= noIterations; j++)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
                IList<Element> elements = collector.OfCategory(bic).ToElements();

                //Start with a 1.5"x1.5"x1.5" cube and gradually add 1/2"
                XYZ min = new XYZ(.021 + point.X - .021 * j, .021 + point.Y - .021 * j, .021 + point.Z - .1 * j);
                XYZ max = new XYZ(.021 + point.X + .021 * j, .021 + point.Y + .021 * j, .021 + point.Z + .1 * j);

                Autodesk.Revit.DB.Outline outline = new Autodesk.Revit.DB.Outline(min, max);
                BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(outline);
                elements = collector.WherePasses(filter).OfCategory(bic).ToElements();

                # region Log
                using (StreamWriter writer2 = new StreamWriter(log, true))
                {
                    writer2.WriteLine("Iteration {0} Finds {1} Possible Mates", j, elements.Count);
                }
                # endregion Log

                if (elements.Count == 0)
                {
                    continue;
                }

                //If we find 1 or more mates, add them to the list
                foreach (Element e in elements)
                {
                    //
                    //Because our bounding box filter starts very small and increases gradually, 
                    //let's assume the first element we find is the one we want
                    //so return that.
                    //
                    elementsFiltered.Add(e);
                    return elementsFiltered;

                }
            }

            return elementsFiltered;
        }


        public static List<Element> FindMatesOfCategory(List<Element> elementsFiltered, BuiltInCategory bic, XYZ point, int noIterations, string systemName, string log, Document doc)
        {
            //
            //Create a list of elements to exclude from the search, such as the current fitting and pipes previously found
            //
            List<ElementId> exclusionList = new List<ElementId>();

            if (elementsFiltered.Count == 1)
            {
                exclusionList.Add(elementsFiltered[0].Id);
            }


            for (int j = 0; j <= noIterations; j++)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
                IList<Element> elements = collector.OfCategory(bic).ToElements();

                //Start with a 1.5"x1.5"x1.5" cube and gradually add 1/2"
                XYZ min = new XYZ(.021 + point.X - .021 * j, .021 + point.Y - .021 * j, .021 + point.Z - .021 * j);
                XYZ max = new XYZ(.021 + point.X + .021 * j, .021 + point.Y + .021 * j, .021 + point.Z + .021 * j);

                Autodesk.Revit.DB.Outline outline = new Autodesk.Revit.DB.Outline(min, max);
                BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(outline);
                elements = collector.WherePasses(filter).Excluding(exclusionList).OfCategory(bic).ToElements();

                # region Log
                using (StreamWriter writer2 = new StreamWriter(log, true))
                {
                    writer2.WriteLine("Iteration {0} Finds {1} Possible Mates", j, elements.Count);
                }
                # endregion Log

                if (elements.Count == 0)
                {
                    continue;
                }

                //If we find 1 or more mates, add them to the list
                foreach (Element e in elements)
                {
                    String mateSystem = e.LookupParameter("Comments").AsString();

                    //
                    //Verify the nearest connector is within 3 inches of the search point.
                    //
                    Connector[] candidateConnectors = MicrodeskHelpers.ConnectorArray(e);
                    Connector candidateNearestConnector = MicrodeskHelpers.NearestConnector(candidateConnectors, point);
                    XYZ candidateDislocation = candidateNearestConnector.Origin - point;

                    # region Log
                    using (StreamWriter writer2 = new StreamWriter(log, true))
                    {
                        writer2.WriteLine("mate ID: {0}", e.Id);
                        writer2.WriteLine("mateSystemAbbr: {0}", mateSystem);
                        writer2.WriteLine("systemAbbreviation: {0}", systemName);
                        writer2.WriteLine("candidateDislocation: {0}", candidateDislocation);
                    }
                    # endregion Log

                    //
                    //A candidate more than 4 inches away is too far
                    //and probably not a good mate
                    //
                    if (candidateDislocation.GetLength() > 0.25)
                    {
                        continue;
                    }

                    if (mateSystem == systemName)
                    {
                        //
                        //Because our bounding box filter starts very small and increases gradually, 
                        //let's assume the first pipe we find is the one we want
                        //so return that.
                        //
                        elementsFiltered.Add(e);
                        return elementsFiltered;
                    }
                }
            }

            return elementsFiltered;
        }


        public static List<Element> FindMatesOfPartType(List<Element> elementsFiltered, PartType partType, XYZ point, int noIterations, string systemName, string log, Document doc)
        {
            string knownMateID = "";

            if (elementsFiltered.Count == 1)
            {
                knownMateID = Convert.ToString(elementsFiltered[0].Id);
            }


            for (int j = 0; j <= noIterations; j++)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
                IList<Element> elements = collector.OfCategory(BuiltInCategory.OST_PipeFitting).ToElements();

                //Start with a 1.5"x1.5"x1.5" cube and gradually add 1/2"
                XYZ min = new XYZ(.021 + point.X - .021 * j, .021 + point.Y - .021 * j, .021 + point.Z - .021 * j);
                XYZ max = new XYZ(.021 + point.X + .021 * j, .021 + point.Y + .021 * j, .021 + point.Z + .021 * j);

                Autodesk.Revit.DB.Outline outline = new Autodesk.Revit.DB.Outline(min, max);
                BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(outline);
                elements = collector.WherePasses(filter).OfCategory(BuiltInCategory.OST_PipeFitting).ToElements();

                # region Log
                using (StreamWriter writer2 = new StreamWriter(log, true))
                {
                    writer2.WriteLine("Iteration {0} Finds {1} Possible Mates", j, elements.Count);
                }
                # endregion Log

                if (elements.Count == 0)
                {
                    continue;
                }

                //If we find 1 or more mates, add them to the list
                foreach (Element e in elements)
                {
                    //
                    //First Verify the Part Type
                    //for example, tee, elbow, etc.
                    //
                    FamilyInstance fi = e as FamilyInstance;

                    MEPModel mM = fi.MEPModel;

                    MechanicalFitting mF = mM as MechanicalFitting;

                    if (mF.PartType != partType)
                    {
                        //
                        //The attached element is not the requested PartType
                        //
                        continue;
                    }

                    String mateSystem = e.LookupParameter("Comments").AsString();

                    //
                    //Verify the nearest connector is within 3 inches of the search point.
                    //
                    Connector[] candidateConnectors = MicrodeskHelpers.ConnectorArray(e);
                    Connector candidateNearestConnector = MicrodeskHelpers.NearestConnector(candidateConnectors, point);
                    XYZ candidateDislocation = candidateNearestConnector.Origin - point;

                    # region Log
                    using (StreamWriter writer2 = new StreamWriter(log, true))
                    {
                        writer2.WriteLine("mate ID: {0}", e.Id);
                        writer2.WriteLine("mateSystemAbbr: {0}", mateSystem);
                        writer2.WriteLine("systemAbbreviation: {0}", systemName);
                        writer2.WriteLine("candidateDislocation: {0}", candidateDislocation);
                    }
                    # endregion Log

                    //
                    //A candidate more than 4 inches away is too far
                    //and probably not a good mate
                    //
                    if (candidateDislocation.GetLength() > 0.25)
                    {
                        continue;
                    }

                    if (mateSystem == systemName)
                    {
                        //
                        //Because our bounding box filter starts very small and increases gradually, 
                        //let's assume the first pipe we find is the one we want
                        //so return that.
                        //
                        elementsFiltered.Add(e);
                        return elementsFiltered;
                    }
                }
            }

            return elementsFiltered;
        }


        public static List<Element> FindMates(List<Element> elementsFiltered, XYZ point, int noIterations, string systemName, string log, Document doc)
        {
            string knownMateID = "";

            //
            //Create a list of elements to exclude from the search, such as the current fitting and pipes previously found
            //
            List<ElementId> exclusionList = new List<ElementId>();

            //
            //If a pipe was found on the previous search, exclude it
            //
            if (elementsFiltered.Count > 0)
            {
                # region Log
                using (StreamWriter writer2 = new StreamWriter(log, true))
                {
                    writer2.WriteLine("Creating Exclusion List");
                }
                # endregion Log

                knownMateID = Convert.ToString(elementsFiltered[0].Id);
                exclusionList.Add(elementsFiltered[0].Id);
            }


            for (int j = 0; j <= noIterations; j++)
            {
                //create a filter for either pipes, pipe fittings or sprinklers
                ElementCategoryFilter pipeFilter = new ElementCategoryFilter(BuiltInCategory.OST_PipeCurves);
                ElementCategoryFilter pipeFittingFilter = new ElementCategoryFilter(BuiltInCategory.OST_PipeFitting);
                ElementCategoryFilter pipeAccessoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory);
                ElementCategoryFilter sprinklerFilter = new ElementCategoryFilter(BuiltInCategory.OST_Sprinklers);

                LogicalOrFilter orFilter1 = new LogicalOrFilter(pipeFilter, pipeFittingFilter);
                LogicalOrFilter orFilter2 = new LogicalOrFilter(orFilter1, pipeAccessoryFilter);
                LogicalOrFilter orFilter = new LogicalOrFilter(orFilter2, sprinklerFilter);

                FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
                IList<Element> elements = collector.WherePasses(orFilter).ToElements();

                //Start with a 1.5"x1.5"x1.5" cube and gradually add 1/2"
                XYZ min = new XYZ(.021 + point.X - .021 * j, .021 + point.Y - .021 * j, .021 + point.Z - .021 * j);
                XYZ max = new XYZ(.021 + point.X + .021 * j, .021 + point.Y + .021 * j, .021 + point.Z + .021 * j);

                Autodesk.Revit.DB.Outline outline = new Autodesk.Revit.DB.Outline(min, max);
                BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(outline);

                if (exclusionList == null || exclusionList.Count == 0)
                {
                    elements = collector.WherePasses(bbFilter).WherePasses(orFilter).ToElements();
                }
                else
                {
                    elements = collector.WherePasses(bbFilter).WherePasses(orFilter).Excluding(exclusionList).ToElements();
                }


                # region Log
                using (StreamWriter writer2 = new StreamWriter(log, true))
                {
                    writer2.WriteLine("Iteration {0} Finds {1} Possible Mates", j, elements.Count);
                }
                # endregion Log

                if (elements.Count == 0)
                {
                    continue;
                }

                //If we find 1 or more mates, add them to the list
                foreach (Element e in elements)
                {
                    String mateSystem = e.LookupParameter("Comments").AsString();

                    //
                    //Verify the nearest connector is within 3 inches of the search point.
                    //
                    Connector[] candidateConnectors = MicrodeskHelpers.ConnectorArray(e);
                    Connector candidateNearestConnector = MicrodeskHelpers.NearestConnector(candidateConnectors, point);
                    XYZ candidateDislocation = candidateNearestConnector.Origin - point;

                    # region Log
                    using (StreamWriter writer2 = new StreamWriter(log, true))
                    {
                        writer2.WriteLine("mate ID: {0}", e.Id);
                        writer2.WriteLine("mateSystemAbbr: {0}", mateSystem);
                        writer2.WriteLine("systemAbbreviation: {0}", systemName);
                        writer2.WriteLine("candidateDislocation: {0}", candidateDislocation);
                    }
                    # endregion Log

                    if (candidateNearestConnector.IsConnected)
                    {
                        # region Log
                        using (StreamWriter writer2 = new StreamWriter(log, true))
                        {
                            writer2.WriteLine("mate {0} is already connected", e.Id);
                        }
                        # endregion Log
                        continue;
                    }

                    if (candidateDislocation.GetLength() > 0.25)
                    {
                        continue;
                    }

                    if (mateSystem == systemName || systemName == null)
                    {
                        elementsFiltered.Add(e);
                        return elementsFiltered;
                    }
                }
            }
            #region Log
            using (StreamWriter writer = new StreamWriter(log, true))
            {
                writer.WriteLine("Nothing Found, returning elementsFiltered");
            }
            #endregion Log
            return elementsFiltered;
        }


        public static List<Element> FindMatesWithCenterPoint(List<Element> elementsFiltered, XYZ point, XYZ centerPoint, int noIterations, string systemName, string log, Document doc)
        {
            //
            //Create a list of elements to exclude from the search, such as the current fitting and pipes previously found
            //
            List<ElementId> exclusionList = new List<ElementId>();

            //
            //If a pipe was found on the previous search, exclude it
            //
            if (elementsFiltered.Count > 0)
            {
                foreach (Element e in elementsFiltered)
                {
                    exclusionList.Add(e.Id);
                }
            }


            for (int j = 0; j <= noIterations; j++)
            {
                //create a filter for either pipes, pipe fittings or sprinklers
                ElementCategoryFilter pipeFilter = new ElementCategoryFilter(BuiltInCategory.OST_PipeCurves);
                ElementCategoryFilter pipeFittingValves = new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory);
                ElementCategoryFilter pipeFittingFilter = new ElementCategoryFilter(BuiltInCategory.OST_PipeFitting);
                ElementCategoryFilter sprinklerFilter = new ElementCategoryFilter(BuiltInCategory.OST_Sprinklers);

                LogicalOrFilter orFilter1 = new LogicalOrFilter(pipeFilter, pipeFittingFilter);
                LogicalOrFilter orFilter2 = new LogicalOrFilter(orFilter1, pipeFittingValves);
                LogicalOrFilter orFilter = new LogicalOrFilter(orFilter2, sprinklerFilter);

                FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
                IList<Element> elements = collector.WherePasses(orFilter).ToElements();

                //Start with a 1.5"x1.5"x1.5" cube and gradually add 1/2"
                XYZ min = new XYZ(.021 + point.X - .021 * j, .021 + point.Y - .021 * j, .021 + point.Z - .021 * j);
                XYZ max = new XYZ(.021 + point.X + .021 * j, .021 + point.Y + .021 * j, .021 + point.Z + .021 * j);

                Autodesk.Revit.DB.Outline outline = new Autodesk.Revit.DB.Outline(min, max);
                BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(outline);

                if (exclusionList == null || exclusionList.Count == 0)
                {
                    elements = collector.WherePasses(bbFilter).WherePasses(orFilter).ToElements();
                }
                else
                {
                    elements = collector.WherePasses(bbFilter).WherePasses(orFilter).Excluding(exclusionList).ToElements();
                }


                # region Log
                using (StreamWriter writer2 = new StreamWriter(log, true))
                {
                    writer2.WriteLine("Iteration {0} Finds {1} Possible Mates", j, elements.Count);
                }
                # endregion Log

                if (elements.Count == 0)
                {
                    continue;
                }

                Element closestMate = null;

                //If we find 1 or more mates, add them to the list
                foreach (Element e in elements)
                {
                    String mateSystem = e.LookupParameter("Comments").AsString();
                    double shortestLineLength = double.MaxValue;

                    //
                    //Verify the nearest connector is within 3 inches of the search point.
                    //
                    Connector[] candidateConnectors = MicrodeskHelpers.ConnectorArray(e);
                    Connector candidateNearestConnector = MicrodeskHelpers.NearestConnector(candidateConnectors, centerPoint);
                    XYZ candidateDislocation = candidateNearestConnector.Origin - point;
                    double candidateCenterDisloc = (candidateNearestConnector.Origin - centerPoint).GetLength();

                    # region Log
                    using (StreamWriter writer2 = new StreamWriter(log, true))
                    {
                        writer2.WriteLine("mate ID: {0}", e.Id);
                        writer2.WriteLine("mateSystemAbbr: {0}", mateSystem);
                        writer2.WriteLine("systemAbbreviation: {0}", systemName);
                        writer2.WriteLine("candidateDislocation: {0}", candidateDislocation);
                        writer2.WriteLine("candidateCenterDisloc: {0}", candidateCenterDisloc);
                    }
                    # endregion Log

                    if (candidateNearestConnector.IsConnected)
                    {
                        # region Log
                        using (StreamWriter writer2 = new StreamWriter(log, true))
                        {
                            writer2.WriteLine("mate {0} is already connected", e.Id);
                        }
                        # endregion Log
                        continue;
                    }

                    if (candidateDislocation.GetLength() > 0.3)
                    {
                        # region Log
                        using (StreamWriter writer2 = new StreamWriter(log, true))
                        {
                            writer2.WriteLine("mate {0} is too far", e.Id);
                        }
                        # endregion Log
                        continue;
                    }

                    //Sprinklers may not have the same system
                    if (mateSystem != systemName && e.Category.Name != "Sprinklers")
                    {
                        # region Log
                        using (StreamWriter writer2 = new StreamWriter(log, true))
                        {
                            writer2.WriteLine("mate {0}'s comments doesn't match the system", e.Id);
                        }
                        # endregion Log
                        continue;
                    }

                    //There could be multiple mates found, so make sure we get the closest
                    if (candidateCenterDisloc < shortestLineLength)
                    {
                        closestMate = e;
                        shortestLineLength = candidateCenterDisloc;
                    }
                }

                if (closestMate == null)
                {
                    continue;
                }

                elementsFiltered.Add(closestMate);
                return elementsFiltered;
            }
            #region Log
            using (StreamWriter writer = new StreamWriter(log, true))
            {
                writer.WriteLine("Nothing Found, returning elementsFiltered");
            }
            #endregion Log
            return elementsFiltered;
        }


        public static List<Element> FindMatesWithBoundingBox(List<Element> elementsFiltered, XYZ point, XYZ centerPoint, double searchBoxWidth, double searchBoxDepth, double searchBoxHeight, int noIterations, string systemName, string log, Document doc, Autodesk.Revit.ApplicationServices.Application app, bool testConnector)
        {
            # region Log
            using (StreamWriter writer2 = new StreamWriter(log, true))
            {
                writer2.WriteLine("FindMatesWithBoundingBox");
            }
            # endregion Log
            //
            //Create a list of elements to exclude from the search, such as the current fitting and pipes previously found
            //
            List<ElementId> exclusionList = new List<ElementId>();
            ElementId previouslyFound1 = null;
            ElementId previouslyFound2 = null;

            //
            //If a pipe was found on the previous search, exclude it
            //
            if (elementsFiltered.Count > 0)
            {

                exclusionList.Add(elementsFiltered[0].Id);
                previouslyFound1 = elementsFiltered[0].Id;
                # region Log
                using (StreamWriter writer2 = new StreamWriter(log, true))
                {
                    writer2.WriteLine("Excluding {0} from search", previouslyFound1);
                }
                # endregion Log
            }
            if (elementsFiltered.Count > 1)
            {

                exclusionList.Add(elementsFiltered[1].Id);
                previouslyFound2 = elementsFiltered[1].Id;
                # region Log
                using (StreamWriter writer2 = new StreamWriter(log, true))
                {
                    writer2.WriteLine("Excluding {0} from search", previouslyFound2);
                }
                # endregion Log
            }

            # region Log
            using (StreamWriter writer2 = new StreamWriter(log, true))
            {
                writer2.WriteLine("Exclusion List Count {0}", exclusionList.Count);
            }
            # endregion Log

            //
            //detremine the max appropriate dislocation a pipe sould be from the center
            //
            double maxBoxWidth = searchBoxWidth;
            if (searchBoxDepth > maxBoxWidth)
            {
                maxBoxWidth = searchBoxDepth;
            }
            if (searchBoxHeight > maxBoxWidth)
            {
                maxBoxWidth = searchBoxHeight;
            }



            for (int j = 0; j <= noIterations; j++)
            {
                # region Log
                using (StreamWriter writer2 = new StreamWriter(log, true))
                {
                    writer2.WriteLine("Iteration {0}", j);
                }
                # endregion Log

                //create a filter for either pipes, pipe fittings or sprinklers
                ElementCategoryFilter pipeFilter = new ElementCategoryFilter(BuiltInCategory.OST_PipeCurves);
                ElementCategoryFilter pipeFittingValves = new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory);
                ElementCategoryFilter pipeFittingFilter = new ElementCategoryFilter(BuiltInCategory.OST_PipeFitting);
                ElementCategoryFilter sprinklerFilter = new ElementCategoryFilter(BuiltInCategory.OST_Sprinklers);


                LogicalOrFilter orFilter1 = new LogicalOrFilter(pipeFilter, pipeFittingFilter);
                LogicalOrFilter orFilter2 = new LogicalOrFilter(orFilter1, pipeFittingValves);
                LogicalOrFilter orFilter = new LogicalOrFilter(orFilter2, sprinklerFilter);

                FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
                IList<Element> elements = collector.WherePasses(orFilter).ToElements();

                //Start with a cube and gradually add 1/2"
                XYZ min = new XYZ(point.X - searchBoxWidth - .021 * j, point.Y - searchBoxDepth - .021 * j, point.Z - searchBoxHeight - .021 * j);
                XYZ max = new XYZ(point.X + searchBoxWidth + .021 * j, point.Y + searchBoxDepth + .021 * j, point.Z + searchBoxHeight + .021 * j);


                //Transaction lineTransaction = new Transaction(doc, "Microdesk.DrawLines");
                //lineTransaction.Start();

                //try
                //{

                //    Element newModelLine1 = MicrodeskHelpers.CreateModelLine(min, max, doc, app);
                //    #region Log
                //    using (StreamWriter writer = new StreamWriter(log, true))
                //    {
                //        writer.WriteLine("Model Lines: {0}", newModelLine1.Id);
                //    }
                //    #endregion Log
                //}
                //catch
                //{
                //    #region Log
                //    using (StreamWriter writer = new StreamWriter(log, true))
                //    {
                //        writer.WriteLine("Model Lines Failed");
                //    }
                //    #endregion Log
                //}
                //lineTransaction.Commit();


                Autodesk.Revit.DB.Outline outline = new Autodesk.Revit.DB.Outline(min, max);
                BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(outline);

                if (exclusionList == null || exclusionList.Count == 0)
                {
                    elements = collector.WherePasses(bbFilter).WherePasses(orFilter).ToElements();
                }
                else
                {
                    # region Log
                    using (StreamWriter writer2 = new StreamWriter(log, true))
                    {
                        writer2.WriteLine("Excluding previous mates");
                    }
                    # endregion Log
                    //elements = collector.WherePasses(bbFilter).WherePasses(orFilter).Excluding(exclusionList).ToElements();
                    elements = collector.WherePasses(bbFilter).WherePasses(orFilter).ToElements();
                }


                # region Log
                using (StreamWriter writer2 = new StreamWriter(log, true))
                {
                    writer2.WriteLine("Iteration {0} Finds {1} Possible Mates", j, elements.Count);
                }
                # endregion Log

                if (elements.Count == 0)
                {
                    continue;
                }

                Element closestMate = null;

                //If we find 1 or more mates, add them to the list
                foreach (Element e in elements)
                {

                    if (previouslyFound1 == e.Id || previouslyFound2 == e.Id)
                    {
                        continue;
                    }


                    String mateSystem = e.LookupParameter("Comments").AsString();
                    double shortestLineLength = double.MaxValue;

                    //
                    //Verify the nearest connector is within 3 inches of the search point.
                    //
                    Connector[] candidateConnectors = MicrodeskHelpers.ConnectorArray(e);
                    Connector candidateNearestConnector = MicrodeskHelpers.NearestConnector(candidateConnectors, centerPoint);
                    XYZ candidateDislocation = candidateNearestConnector.Origin - point;
                    double candidateCenterDisloc = (candidateNearestConnector.Origin - centerPoint).GetLength() - maxBoxWidth;

                    # region Log
                    using (StreamWriter writer2 = new StreamWriter(log, true))
                    {
                        writer2.WriteLine("mate ID: {0}", e.Id);
                        writer2.WriteLine("mateSystem: {0}", mateSystem);
                        writer2.WriteLine("systemName: {0}", systemName);
                        writer2.WriteLine("candidateDislocation: {0}", candidateDislocation);
                        writer2.WriteLine("candidateCenterDisloc: {0}", candidateCenterDisloc);
                    }
                    # endregion Log

                    if (testConnector)
                    {
                        if (candidateNearestConnector.IsConnected)
                        {
                            # region Log
                            using (StreamWriter writer2 = new StreamWriter(log, true))
                            {
                                writer2.WriteLine("mate {0} is already connected", e.Id);
                            }
                            # endregion Log
                            continue;
                        }


                        if (candidateDislocation.GetLength() > 0.5)
                        {
                            # region Log
                            using (StreamWriter writer2 = new StreamWriter(log, true))
                            {
                                writer2.WriteLine("mate {0} is too far", e.Id);
                            }
                            # endregion Log
                            continue;
                        }
                    }

                    //Sprinklers may not have the same system
                    if (mateSystem != systemName &&
                        !mateSystem.ToUpper().Contains("BYPASS") &&
                        !systemName.ToUpper().Contains("BYPASS") &&
                        !systemName.ToUpper().Contains("VALVE") &&
                        e.Category.Name != "Sprinklers")
                    {
                        # region Log
                        using (StreamWriter writer2 = new StreamWriter(log, true))
                        {
                            writer2.WriteLine("mate {0}'s comments doesn't match the system", e.Id);
                        }
                        # endregion Log
                        continue;
                    }

                    //There could be multiple mates found, so make sure we get the closest
                    if (candidateCenterDisloc < shortestLineLength)
                    {
                        closestMate = e;
                        shortestLineLength = candidateCenterDisloc;
                    }
                }

                if (closestMate == null)
                {
                    continue;
                }

                elementsFiltered.Add(closestMate);
                return elementsFiltered;
            }
            #region Log
            using (StreamWriter writer = new StreamWriter(log, true))
            {
                writer.WriteLine("Nothing Found, returning elementsFiltered");
            }
            #endregion Log
            return elementsFiltered;
        }


        public static Element CreateModelLine(XYZ p, XYZ q, Document doc, Autodesk.Revit.ApplicationServices.Application app)
        {
            if (p.IsAlmostEqualTo(q))
            {
                throw new ArgumentException(
                  "Expected two different points.");
            }

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(p, q);
            if (null == line)
            {
                throw new Exception(
                  "Geometry line creation failed.");
            }

            Element newModelLine = doc.Create.NewModelCurve(line, MicrodeskHelpers.NewSketchPlanePassLine(line, doc, app));

            return newModelLine;
        }


        public static SketchPlane NewSketchPlanePassLine(Autodesk.Revit.DB.Line line, Document doc, Autodesk.Revit.ApplicationServices.Application app)
        {
            XYZ p = line.GetEndPoint(0);
            XYZ q = line.GetEndPoint(1);
            XYZ norm;
            /*
            if (p.X == q.X)
            {
                norm = XYZ.BasisX;
            }
            else if (p.Y == q.Y)
            {
                norm = XYZ.BasisY;
            }
            else if (p.Z == q.Z)
            {
                norm = XYZ.BasisZ;
            }
            else
            {
                norm = p.CrossProduct(q);
            }
            */
            norm = p.CrossProduct(q);

#if (R2016)
            Plane plane = app.Create.NewPlane(norm, p);
#else
            Plane plane = Plane.CreateByNormalAndOrigin(norm, p);
#endif



            return SketchPlane.Create(doc, plane);
        }


        public static SketchPlane NewSketchPlaneThreePoints(XYZ p, XYZ q, XYZ r, Document doc, Autodesk.Revit.ApplicationServices.Application app)
        {
            XYZ pq = q - p;
            XYZ pr = r - p;

            XYZ norm = pq.CrossProduct(pr);

#if (R2016)
            Plane plane = app.Create.NewPlane(norm, p);
#else
            Plane plane = Plane.CreateByNormalAndOrigin(norm, p);
#endif

            return SketchPlane.Create(doc, plane);
        }

        //
        //Revit Families
        //

        public static List<FileInfo> GetFamilyFiles(DirectoryInfo rootDirInfo, List<FileInfo> allFiles)
        {

            DirectoryInfo[] subDirectories = null;
            FileInfo[] files = null;


            try
            {
                files = rootDirInfo.GetFiles("*.rfa");
            }
            catch { }

            foreach (FileInfo fI in files)
            {
                allFiles.Add(fI);
            }


            subDirectories = rootDirInfo.GetDirectories();

            foreach (System.IO.DirectoryInfo dirInfo in subDirectories)
            {
                allFiles = GetFamilyFiles(dirInfo, allFiles);
            }

            return allFiles;
        }


        public static bool CheckWorkset(Document doc, string name)
        {
            FilteredWorksetCollector collector = new FilteredWorksetCollector(doc);

            collector.OfKind(WorksetKind.UserWorkset);
            IList<Workset> worksets = collector.ToWorksets();

            foreach (Workset workset in worksets)
            {
                if (workset.Name == name)
                {
                    return true;
                }
            }

            return false;
        }


        public static Workset GetCreateWorkset(Document doc, string name)
        {
            FilteredWorksetCollector collector = new FilteredWorksetCollector(doc);

            collector.OfKind(WorksetKind.UserWorkset);
            IList<Workset> worksets = collector.ToWorksets();

            foreach (Workset workset in worksets)
            {
                if (workset.Name == name)
                {
                    return workset;
                }
            }

            Transaction trans = new Transaction(doc, "Create Workset");
            trans.Start();
            Workset newWorkset = Workset.Create(doc, name);

            //
            // Get the workset’s default visibility    
            //
            WorksetDefaultVisibilitySettings defaultVisibility = WorksetDefaultVisibilitySettings.GetWorksetDefaultVisibilitySettings(doc);

            //
            // Set it to 'false'
            //
            if (true == defaultVisibility.IsWorksetVisible(newWorkset.Id))
            {
                defaultVisibility.SetWorksetVisibility(newWorkset.Id, false);
            }

            //
            // Set the Workset visible in the active View
            //
            doc.ActiveView.SetWorksetVisibility(newWorkset.Id, WorksetVisibility.Visible);

            trans.Commit();
            return newWorkset;
        }


        public static Category CreateSubCategory(Document doc, string subCatergoryName)
        {
            // create a new subcategory 
            Category cat = doc.OwnerFamily.FamilyCategory;
            Category subCat = doc.Settings.Categories.NewSubcategory(cat, subCatergoryName);

            return subCat;
        }

        //Retrieve the group from a shared parameter file, or create it if it doesn't exist
        public static DefinitionGroup GetGroup(DefinitionFile myDefinitionFile, string groupName, string log)
        {
            DefinitionGroup defGroup = null;

            //Get the existing groups
            DefinitionGroups existingGroups = myDefinitionFile.Groups;

            //Iterate though the groups and return the desired group if it's found
            foreach (DefinitionGroup dG in existingGroups)
            {
                if (dG.Name == groupName)
                {
                    #region Log
                    using (StreamWriter writer = new StreamWriter(log, true))
                    {
                        writer.WriteLine("Group {0} already exists", groupName);
                    }
                    #endregion Log
                    defGroup = dG;

                    return defGroup;
                }
            }

            //The desired group wasn't found at this point, so create a new one
            #region Log
            using (StreamWriter writer = new StreamWriter(log, true))
            {
                writer.WriteLine("Group {0} not found", groupName);
            }
            #endregion Log
            defGroup = existingGroups.Create(groupName);
            #region Log
            using (StreamWriter writer = new StreamWriter(log, true))
            {
                writer.WriteLine("Group {0} created", defGroup.Name);
            }
            #endregion Log

            return defGroup;
        }


        //Retrieve the parameter definition from a shared parameter file, or create it if it doesn't exist
        public static ExternalDefinition GetDefinition(DefinitionFile sharedParamFile, string definitionName, ParameterType paramType, ref Guid guid, DefinitionGroup group, string log)
        {
            List<ExternalDefinition> exDefs = new List<ExternalDefinition>();
            ExternalDefinition exDef = null;

            //#region Log
            //using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            //{
            //    writer.WriteLine("GetDefinition");
            //}
            //#endregion Log

            foreach (DefinitionGroup myGroup in sharedParamFile.Groups)
            {
                foreach (ExternalDefinition definition in myGroup.Definitions)
                {
                    if (definition.Name == definitionName)
                    {
                        exDef = definition;
                        return exDef;
                    }
                }
            }


            ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(definitionName, paramType);
            exDef = group.Definitions.Create(options) as ExternalDefinition;
            //Definition newDef = group.Definitions.Create(definitionName, paramType, true, ref guid);

            return exDef as ExternalDefinition;
        }


        //
        //EXCEL
        //

        public static string OpenExcelFile()
        {
            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
            //openFileDialog.InitialDirectory = @"C:\";
            openFileDialog.Filter = "Excel Files (.xlsx)|*.xlsx|(.xls)|*.xls";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = false;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
            }

            return openFileDialog.FileName;
        }


        public static string[] OpenMultipleFiles()
        {
            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.Multiselect = true;
            //openFileDialog.InitialDirectory = @"C:\";
            openFileDialog.Filter = "Excel Files (.dwg)|*.dwg";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = false;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
            }

            return openFileDialog.FileNames;
        }


        public static string GetExcelFile()
        {
            string xlFile = @"C:\default";
            System.Windows.Forms.SaveFileDialog saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog1.Filter = "Excel File|*.xls";
            saveFileDialog1.Title = "Save an Excel File";
            saveFileDialog1.ShowDialog();

            if (saveFileDialog1.FileName != "")
            {
                xlFile = saveFileDialog1.FileName;
            }

            return xlFile;
        }


        public static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        /// <summary>
        /// <para>Specify a local function before use. Example: </para>
        /// <para>Action revitAction = () => MethodThatModifyRevitDB(...);</para>
        /// <para>document.DoTransaction(transactionName, revitAction, out Exception transactionException);</para>
        /// </summary>
        /// <param name="document"></param>
        /// <param name="transactionName"></param>
        /// <param name="revitAction"></param>
        /// <param name="transactionException"></param>
        /// <returns></returns>
        public static bool DoTransaction(this Document document, string transactionName, System.Action revitAction, out Exception transactionException)
        {
            transactionException = null;
            using (Transaction transaction = new Transaction(document, transactionName))
            {
                transaction.Start();
                try
                {
                    revitAction();

                    transaction.Commit();
                }
                catch (Exception e)
                {
                    transaction.RollBack();
                    transactionException = e;
                }
            }

            return transactionException == null;
        }
    }
}
