using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;

using System.Windows.Forms;
using System.Diagnostics;
using System.Collections;
using System.Xml;
using System.IO;

using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;


namespace NBK.Spool
{
    public class LevelManager
    {
        protected Autodesk.Revit.ApplicationServices.Application m_application;
        protected Document m_document;

        public IList<RevitElement> m_levels = new List<RevitElement>();
        public int m_topIndex = 0;
        public int m_bottomIndex = 0;
        public double maxZ = 999;
        public double minZ = -999;



        public LevelManager(Autodesk.Revit.ApplicationServices.Application app, Document doc)
        {
            m_application = app;
            m_document = doc;
            GetLevels();
        }

        public IList<RevitElement> Levels
        {
            get
            {
                return m_levels;
            }
        }

        public int TopIndex
        {
            get
            {
                return m_topIndex;
            }
        }

        public int BottomIndex
        {
            get
            {
                return m_bottomIndex;
            }
        }

        public double MinZ
        {
            get
            {
                return minZ;
            }
        }

        public double MaxZ
        {
            get
            {
                return maxZ;
            }
        }

        public void GetLevels()
        {
            FilteredElementCollector collector = new FilteredElementCollector(m_document).OfClass(typeof(Level));

            foreach (Level level in collector)
            {
                m_levels.Add(new RevitElement(level.Name, level.Elevation));
            }
            m_levels.Add(new RevitElement("Unlimited Down", -9999));
            m_levels.Add(new RevitElement("Unlimited Up", 9999));

            m_levels = m_levels.OrderBy(w => w.Elevation).ToList();

            //
            // Guess the top and bottom
            //
            View3D view3D = m_document.ActiveView as View3D;
            if (view3D != null)
            {
                BoundingBoxXYZ bb = view3D.GetSectionBox();
                minZ = bb.Min.Z + bb.Transform.Origin.Z;
                maxZ = bb.Max.Z + bb.Transform.Origin.Z;

                for (int i = 0; i < m_levels.Count; i++)
                {
                    //
                    //Working our way up through the levels,
                    //once we find a level higher than the bottom of the section box,
                    //test if its the closest level to the bottom of the section box,
                    //or if the level lower is closer
                    if (m_levels[i].Elevation > minZ)
                    {
                        if (Math.Abs(m_levels[i].Elevation - minZ) < Math.Abs(m_levels[i - 1].Elevation - minZ))
                            m_bottomIndex = i;
                        else
                            m_bottomIndex = i - 1;

                        break;
                    }
                    m_bottomIndex = 0;
                }
                for (int i = 0; i < m_levels.Count; i++)
                {
                    if (m_levels[i].Elevation > maxZ)
                    {
                        if (Math.Abs(m_levels[i].Elevation - maxZ) < Math.Abs(m_levels[i - 1].Elevation - maxZ))
                            m_topIndex = i;
                        else
                            m_topIndex = i - 1;

                        break;
                    }
                }
            }

            ViewPlan plan = m_document.ActiveView as ViewPlan;
            if (plan != null)
            {
                Level level = plan.GenLevel;
                PlanViewRange range = plan.GetViewRange();

                ElementId topClipPlane = range.GetLevelId(PlanViewPlane.TopClipPlane);
                Level top = m_document.GetElement(topClipPlane) as Level;
                double offset = range.GetOffset(PlanViewPlane.TopClipPlane);

                minZ = level.Elevation;
                maxZ = top.Elevation + offset;

                for (int i = 0; i < m_levels.Count; i++)
                {
                    if (m_levels[i].Name == level.Name)
                    {
                        m_bottomIndex = i;
                        break;
                    }
                }

                for (int i = 0; i < m_levels.Count; i++)
                {
                    if (m_levels[i].Elevation > maxZ)
                    {
                        if (Math.Abs(m_levels[i].Elevation - maxZ) < Math.Abs(m_levels[i - 1].Elevation - maxZ))
                            m_topIndex = i;
                        else
                            m_topIndex = i - 1;

                        break;
                    }
                }
            }
        }
    }
}
