using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using Autodesk.Revit;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;


using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI.Selection;

namespace NBK.Spool
{
    class ExportContext : IExportContext
    {
        public List<RevitElement> CollectedCurves = new List<RevitElement>();
        private List<Document> m_documents = new List<Document>();
        private List<Document> all_documents = new List<Document>();
        private Autodesk.Revit.ApplicationServices.Application m_application = null;
        private Document m_doc = null;
        private StreamWriter m_writer = null;

        /// <summary>
        /// Constructor
        /// </summary>
        public ExportContext(Application application, Document activeDocument, List<Document> openDocuments)
        {
            m_application = application;
            m_doc = activeDocument;
            m_documents = openDocuments;
            all_documents = openDocuments;
            all_documents.Add(activeDocument);
        }

        public bool Start()
        {
            return true;
        }

        public void Finish()
        {
        }

        public bool IsCanceled()
        {
            return false;
        }

        public RenderNodeAction OnViewBegin(ViewNode node)
        {
            Element view = m_doc.GetElement(node.ViewId);

            //if (view != null)
            //{
            //    using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            //    {
            //        writer.WriteLine("View Name,{0}", view.Name);
            //    }
            //}

            return RenderNodeAction.Proceed;
        }

        public void OnViewEnd(ElementId elementId)
        {
        }

        public RenderNodeAction OnElementBegin(ElementId elementId)
        {
            string elementName = "";
            string linkPath = "";
            string catName = "";
            double height = 1;
            double width = 1;


            foreach (Document doc in all_documents)
            {
                Element e = doc.GetElement(elementId);
                if (e == null)
                {
                    continue;
                }


                try
                {
                    //
                    // If the element ID is found, but it's not a desired category, break the loop
                    //
                    if (e.Category.Id.IntegerValue != (int)BuiltInCategory.OST_CableTray
                    && e.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Conduit
                    && e.Category.Id.IntegerValue != (int)BuiltInCategory.OST_DuctCurves
                    && e.Category.Id.IntegerValue != (int)BuiltInCategory.OST_PipeCurves)
                    {
                        continue;
                    }
                }
                catch
                {
                    // Some elements don't have a category, so we need to catch those exceptions
                    continue;
                }

                //
                // Get information about the curve family
                //
                elementName = e.Name;
                linkPath = doc.PathName;
                catName = e.Category.Name;


                Connector[] connectors = MicrodeskHelpers.ConnectorArray(e);
                if (connectors[0].Shape == ConnectorProfileType.Round)
                {
                    height = connectors[0].Radius * 2;
                    width = connectors[0].Radius * 2;
                }
                else
                {
                    height = connectors[0].Height;
                    width = connectors[0].Width;
                }

                XYZ[] bbCorners = MicrodeskHelpers.CalcBoundingBox(m_application, m_doc, connectors, height, width);

                RevitElement mepCurve =
                    new RevitElement(elementId,
                    e.Name,
                    e.Category.Name,
                    doc,
                    connectors,
                    bbCorners[0],
                    bbCorners[1],
                    height,
                    width);

                CollectedCurves.Add(mepCurve);

                //using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                //{
                //    writer.WriteLine("Element,{0},{1},{2},{3},{4},{5}", 
                //        linkPath, 
                //        elementId.IntegerValue.ToString(), 
                //        mepCurve.CategoryName,
                //        mepCurve.Name,
                //        mepCurve.Height,
                //        mepCurve.Width);
                //}
                break;
            }

            return RenderNodeAction.Proceed;
        }

        public void OnElementEnd(ElementId elementId)
        {

        }

        public RenderNodeAction OnInstanceBegin(InstanceNode node)
        {
            return RenderNodeAction.Proceed;
        }

        public void OnInstanceEnd(InstanceNode node)
        {
            //string instanceName = "";
            //try
            //{
            //    instanceName = node.NodeName;
            //}
            //catch { }
        }

        public RenderNodeAction OnLinkBegin(LinkNode node)
        {
            return RenderNodeAction.Proceed;
        }

        public void OnLinkEnd(LinkNode node)
        {
        }

        public RenderNodeAction OnFaceBegin(FaceNode node)
        {
            return RenderNodeAction.Proceed;
        }

        public void OnFaceEnd(FaceNode node)
        {

        }

        public void OnLight(LightNode node)
        {
            return;
        }

        public void OnRPC(RPCNode node)
        {
            return;
        }

        public void OnMaterial(MaterialNode node)
        {
            return;
        }

        public void OnPolymesh(PolymeshTopology node)
        {
            return;
        }

#if (R2016)
        public void OnDaylightPortal(DaylightPortalNode node)
        {
            return;
        }
#endif


    }
}
