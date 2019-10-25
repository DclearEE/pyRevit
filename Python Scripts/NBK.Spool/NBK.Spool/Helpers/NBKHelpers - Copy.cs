using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using XYZ = Autodesk.Revit.DB.XYZ;

using Reference = Autodesk.Revit.DB.Reference;

using Excel = Microsoft.Office.Interop.Excel;

namespace NBK.Spool
{
    public static class NBKHelpersCopy
    {
        public static List<SpoolNumber> GetExistingSpoolNumbers(Document doc)
        {
            //List<string> existingSpools = new List<string>();

            List<ElementFilter> catFilters = new List<ElementFilter>();
            catFilters.Add(new ElementCategoryFilter(BuiltInCategory.OST_FabricationContainment));
            //catFilters.Add(new ElementCategoryFilter(BuiltInCategory.OST_FabricationDuctwork));
            //catFilters.Add(new ElementCategoryFilter(BuiltInCategory.OST_FabricationHangers));
            catFilters.Add(new ElementCategoryFilter(BuiltInCategory.OST_FabricationPipework));
            ElementFilter orFilter = new LogicalOrFilter(catFilters);

            FilteredElementCollector collector = new FilteredElementCollector(doc).WherePasses(orFilter).OfClass(typeof(FabricationPart));

            List<string> existingSpools = collector
                .Select(w => w.LookupParameter(AppKeys.SpoolParameterName).AsString())
                .Distinct()
                .ToList();

            existingSpools.Sort();

            List<SpoolNumber> spoolNos = new List<SpoolNumber>();
            foreach(string existingSpool in existingSpools)
            {
                if (existingSpool == null || existingSpool == "") continue;

                spoolNos.Add(new SpoolNumber(existingSpool));
            }

            return spoolNos;
        }


        public static List<RevitElement> GetElementsFromReferences(Document doc, IList<Reference> selectedReferences)
        {
            List<RevitElement> selectedElements = new List<RevitElement>();

            FabricationConfiguration fabconfig = FabricationConfiguration.GetFabricationConfiguration(doc);

            foreach (Reference reference in selectedReferences)
            {
                Element element = doc.GetElement(reference);

                // Skip selected elements that are not pipework
                if (element.Category.Id.IntegerValue != (int)BuiltInCategory.OST_FabricationPipework &&
                    element.Category.Id.IntegerValue != (int)BuiltInCategory.OST_PipeAccessory) continue;

                if (element.LookupParameter(AppKeys.SpoolParameterName).AsString() != null
                    && element.LookupParameter(AppKeys.SpoolParameterName).AsString() != "") continue;

                FabricationPart fabPart = element as FabricationPart;
                FabricationPartType fabType = doc.GetElement(element.GetTypeId()) as FabricationPartType;
                string familyTypeName = fabType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString();
                string familyName = fabType.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString();

                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                {
                    writer.WriteLine("reviewing: {0}, {1}", familyTypeName, familyName);
                }

                string length = "";
                Parameter lengthParam = element.get_Parameter(BuiltInParameter.FABRICATION_PART_LENGTH);
                if (lengthParam != null)
                    length = lengthParam.AsValueString();

                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                {
                    writer.WriteLine("length: {0}", length);
                }

                string size = "";
                Parameter sizeParam = element.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE);
                if (sizeParam != null)
                    size = sizeParam.AsString();

                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                {
                    writer.WriteLine("size: {0}", size);
                }

                BoundingBoxXYZ bb = element.get_BoundingBox(doc.ActiveView);

                //Get the directions of all the elements
                List<XYZ> directions = GetPartDirections(element);

                //Classify the part as a pipe or fitting
                string partType = "Pipe Fitting";
                string connector0Name = "";
                string connector1Name = "";

                if (fabPart.IsAStraight())
                {
                    partType = "Pipe";

                    //Get the namee of the connectors
                    Connector[] connectorarray = MicrodeskHelpers.ConnectorArray(element);
                    List<Connector> connectorList = new List<Connector>();

                    // It's possible pipe will have more than 2 connectors if it has a tap
                    // We want to ignore tap connectors
                    foreach (Connector c in connectorarray)
                    {
                        if (c.ConnectorType == ConnectorType.End)
                            connectorList.Add(c);
                    }

                    connector0Name = fabconfig.GetFabricationConnectorName(connectorList[0].GetFabricationConnectorInfo().BodyConnectorId);
                    connector1Name = fabconfig.GetFabricationConnectorName(connectorList[1].GetFabricationConnectorInfo().BodyConnectorId);
                }
                

                RevitElement revitElement = new RevitElement(
                    reference, 
                    element, 
                    familyName, 
                    familyTypeName,
                    partType,
                    size, 
                    length, 
                    bb,
                    directions,
                    connector0Name,
                    connector1Name);

                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                {
                    writer.WriteLine("revitElement: {0}, {1}, {2}, {3}", revitElement.FamilyName, revitElement.FamilyTypeName, revitElement.Size, revitElement.FamilySizeandLength);
                }

                selectedElements.Add(revitElement);
            }

            //List<Element> selectedElements = selectedReferences.Select(x => doc.GetElement(x)).ToList();
            return selectedElements;
        }

        public static List<RevitElement> GetElementsFromIds(Document doc, ICollection<ElementId> selectedIds)
        {

            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("GetElementsFromIds: {0} ids selected", selectedIds.Count);
            }


            List<RevitElement> selectedElements = new List<RevitElement>();

            FabricationConfiguration fabconfig = FabricationConfiguration.GetFabricationConfiguration(doc);

            foreach (ElementId eId in selectedIds)
            {
                Element element = doc.GetElement(eId);

                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                {
                    writer.WriteLine("{0}, {1}", eId.ToString(), element.Name);
                }

                // Skip selected elements that are not pipework
                if (element.Category.Id.IntegerValue != (int)BuiltInCategory.OST_FabricationPipework &&
                    element.Category.Id.IntegerValue != (int)BuiltInCategory.OST_PipeAccessory) continue;

                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                {
                    writer.WriteLine("{0}", element.LookupParameter(AppKeys.SpoolParameterName).AsString());
                }

                if (element.LookupParameter(AppKeys.SpoolParameterName).AsString() != null
                    && element.LookupParameter(AppKeys.SpoolParameterName).AsString() != "") continue;

                Reference referen = new Reference(element);
                FabricationPart fabPart = element as FabricationPart;
                FabricationPartType fabType = doc.GetElement(element.GetTypeId()) as FabricationPartType;
                string familyTypeName = fabType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString();
                string familyName = fabType.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString();

                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                {
                    writer.WriteLine("reviewing: {0}, {1}", familyTypeName, familyName);
                }

                string length = "";
                Parameter lengthParam = element.get_Parameter(BuiltInParameter.FABRICATION_PART_LENGTH);
                if (lengthParam != null)
                    length = lengthParam.AsValueString();

                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                {
                    writer.WriteLine("length: {0}", length);
                }

                string size = "";
                Parameter sizeParam = element.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE);
                if (sizeParam != null)
                    size = sizeParam.AsString();

                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                {
                    writer.WriteLine("size: {0}", size);
                }

                BoundingBoxXYZ bb = element.get_BoundingBox(doc.ActiveView);

                //Get the directions of all the elements
                List<XYZ> directions = GetPartDirections(element);

                //Classify the part as a pipe or fitting
                string partType = "Pipe Fitting";
                string connector0Name = "";
                string connector1Name = "";

                if (fabPart.IsAStraight())
                {
                    partType = "Pipe";

                    //Get the namee of the connectors
                    Connector[] connectorarray = MicrodeskHelpers.ConnectorArray(element);
                    List<Connector> connectorList = new List<Connector>();

                    // It's possible pipe will have more than 2 connectors if it has a tap
                    // We want to ignore tap connectors
                    foreach (Connector c in connectorarray)
                    {
                        if (c.ConnectorType == ConnectorType.End)
                            connectorList.Add(c);
                    }

                    connector0Name = fabconfig.GetFabricationConnectorName(connectorList[0].GetFabricationConnectorInfo().BodyConnectorId);
                    connector1Name = fabconfig.GetFabricationConnectorName(connectorList[1].GetFabricationConnectorInfo().BodyConnectorId);
                }


                RevitElement revitElement = new RevitElement(
                    referen,
                    element,
                    familyName,
                    familyTypeName,
                    partType,
                    size,
                    length,
                    bb,
                    directions,
                    connector0Name,
                    connector1Name);

                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                {
                    writer.WriteLine("revitElement: {0}, {1}, {2}, {3}", revitElement.FamilyName, revitElement.FamilyTypeName, revitElement.Size, revitElement.FamilySizeandLength);
                }

                selectedElements.Add(revitElement);
            }

            //List<Element> selectedElements = selectedReferences.Select(x => doc.GetElement(x)).ToList();
            return selectedElements;
        }

        public static List<XYZ> GetPartDirections(Element element)
        {
            ConnectorSet cS = null;
            List<XYZ> directions = new List<XYZ>();

            FabricationPart fabPart = element as FabricationPart;

            if (fabPart == null)
                return null;

            cS = fabPart.ConnectorManager.Connectors;

            foreach(Connector c in cS)
            {
                XYZ unit = c.CoordinateSystem.BasisZ;

                //for now, ignore vertical connectors
                if (Math.Round(unit.X, 2) == 0 && 
                    Math.Round(unit.Y, 2) == 0)
                    continue;

                XYZ flatVector = (new XYZ(Math.Abs(unit.X), Math.Abs(unit.Y), 0)).Normalize();

                directions.Add(flatVector);
            }

            return directions;
        }


        public static void NumberSpoolElements(List<RevitElement> spoolElements)
        {
            var queryParts =
                    from revitElement in spoolElements
                    group revitElement by revitElement.FamilySizeandLength into newGroup
                    orderby newGroup.Key
                    select newGroup;

            int i = 1;
            foreach (var nameGroup in queryParts)
            {
                foreach(RevitElement part in nameGroup)
                {
                    part.Element.LookupParameter(AppKeys.SpoolItemParameterName).Set(i.ToString());
                }

                i++;
            }
        }


        public static BoundingBoxXYZ GetBoudingBox(Document doc, List<RevitElement> selectedElements)
        {
            double minX = 99999;
            double minY = 99999;
            double minZ = 99999;
            double maxX = -99999;
            double maxY = -99999;
            double maxZ = -99999;

            foreach (RevitElement e in selectedElements)
            {
                if (e.Min.X < minX)
                    minX = e.Min.X;

                if (e.Min.Y < minY)
                    minY = e.Min.Y;

                if (e.Min.Z < minZ)
                    minZ = e.Min.Z;

                if (e.Max.X > maxX)
                    maxX = e.Max.X;

                if (e.Max.Y > maxY)
                    maxY = e.Max.Y;

                if (e.Max.Z > maxZ)
                    maxZ = e.Max.Z;
            }

            BoundingBoxXYZ bbNew = new BoundingBoxXYZ();
            bbNew.Min = new XYZ(minX - .25, minY - .25, minZ - .25);
            bbNew.Max = new XYZ(maxX + .25, maxY + .25, maxZ + .25);

            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("View3D Max: {0}, {1}, {2}", Math.Round(maxX, 3), Math.Round(maxY, 3), Math.Round(maxZ, 3));
                writer.WriteLine("View3D Min: {0}, {1}, {2}", Math.Round(minX, 3), Math.Round(minY, 3), Math.Round(minZ, 3));
            }

            return bbNew;
        }


        public static int GetScale(BoundingBoxXYZ bb)
        {
            // 1" = 1'     ... 12
            // 3/4" = 1'   ... 18
            // 1/2" = 1'   ... 24
            // 1/4" = 1'   ... 48
            // 1/8" = 1'   ... 96

            int scale = 48;




            return scale;
        }


        public static ElementId GetViewTypeId(Document doc, string viewName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType));
            foreach(ViewFamilyType v in collector)
            {
                if(v.Name == viewName)
                    return v.Id;
            }

            return null;
        }


        public static ElementId GetViewTemplate(Document doc, string filterName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));

            List<Element> viewTemplates = collector.Where(w => w.Name == filterName).ToList();

            if(viewTemplates.Count == 0)
            {
                TaskDialog.Show("Warning", 
                    "The active model does not contain a schedule view template named "
                    + filterName + 
                    ", which this command requires");
            }

            Element viewTemplate = viewTemplates.First();

            return viewTemplate.Id;

        }


        public static void AddSpoolNumberToElements(Document doc, List<RevitElement> selectedElements, string spoolNumber)
        {
            foreach(RevitElement e in selectedElements)
            {
                doc.GetElement(e.ElementId).LookupParameter(AppKeys.SpoolParameterName).Set(spoolNumber);
                doc.GetElement(e.ElementId).LookupParameter(AppKeys.SpoolTypeParameterName).Set(e.PartType);
                doc.GetElement(e.ElementId).LookupParameter(AppKeys.SpoolConn0ParameterName).Set(e.Connector0Name);
                doc.GetElement(e.ElementId).LookupParameter(AppKeys.SpoolConn1ParameterName).Set(e.Connector1Name);
            }
        }


        public static ElementId GetTypeId(Document doc, BuiltInCategory bic, Type type, string name)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc).OfCategory(bic).OfClass(type);

            foreach(Element e in collector)
            {
                if (e.Name == name)
                    return e.Id;
            }
            return null;
        }


        public static ParameterFilterElement CreateFilter(Document doc, string value, bool conditionEquals)
        {
            string filterName = "";

            if (conditionEquals)
                filterName = "Spool " + value;
            else
                filterName = "Isolate Spool " + value;

            ParameterFilterElement pfe = null;

            FilteredElementCollector coll = new FilteredElementCollector(doc).OfClass(typeof(ParameterFilterElement));
            IEnumerable<Element> paramFilters = from element in coll where element.Name.Equals(filterName) select element;
            foreach (Element element in paramFilters)
            {
                if (element.Name.Equals(filterName))
                {
                    pfe = element as ParameterFilterElement;
                }
            }

            if (pfe == null)
            {
                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                {
                    writer.WriteLine("ParameterFilterElement pfe is null");
                }

                List<ElementId> categories = new List<ElementId>();

                categories.Add(new ElementId(BuiltInCategory.OST_PipeAccessory));
                categories.Add(new ElementId(BuiltInCategory.OST_FabricationDuctwork));
                categories.Add(new ElementId(BuiltInCategory.OST_FabricationPipework));
                categories.Add(new ElementId(BuiltInCategory.OST_GenericModel));

                pfe = ParameterFilterElement.Create(doc, filterName, categories);
            }

            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("ParameterFilterElement pfe: {0}", pfe.Id);
            }

            FilteredElementCollector parameterCollector = new FilteredElementCollector(doc);
            Parameter parameter = parameterCollector.OfClass(typeof(FabricationPart)).FirstElement().LookupParameter(AppKeys.SpoolParameterName);

            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("parameter: {0}", parameter.ToString());
            }

            List<FilterRule> filterRules = new List<FilterRule>();

            if(conditionEquals)
                filterRules.Add(ParameterFilterRuleFactory.CreateEqualsRule(parameter.Id, value, true));
            else
                filterRules.Add(ParameterFilterRuleFactory.CreateNotEqualsRule(parameter.Id, value, true));

            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("filterRules: {0}", filterRules.Count);
            }

#if R2020
            ElementParameterFilter filter = new ElementParameterFilter(filterRules);

            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("filter: {0}", filter.ToString());
            }

            pfe.SetElementFilter(filter);
#else
            pfe.SetRules(filterRules);
#endif
            return pfe;
        }


        public static void AddFiltertoView(Document doc, Autodesk.Revit.DB.View view, ParameterFilterElement pfe)
        {
            OverrideGraphicSettings filterSettings = new OverrideGraphicSettings();

            //filterSettings.SetCutFillColor(colo);
            //filterSettings.SetCutFillPatternId(xx);
            view.SetFilterVisibility(pfe.Id, false);
            view.SetFilterOverrides(pfe.Id, filterSettings);
        }


        public static List<Element> GetViewportTypes(Document doc)
        {
            FilteredElementCollector collectorTypes = new FilteredElementCollector(doc).OfClass(typeof(ElementType));
            List<Element> viewportTypes = new List<Element>();
            foreach (Element elemType in collectorTypes)
            {
                //
                // Viewport Types would be the only ElementType using the following parameter
                //
                if (elemType.get_Parameter(BuiltInParameter.VIEWPORT_ATTR_SHOW_EXTENSION_LINE) != null)
                    viewportTypes.Add(elemType);
            }

            return viewportTypes;
        }


        public static void SetViewportType(Document doc, Viewport viewport, List<Element> viewportTypes, string viewportTypeName)
        {
            foreach(Element viewportType in viewportTypes)
            {
                if(viewportType.Name == viewportTypeName)
                {
                    viewport.ChangeTypeId(viewportType.Id);
                    return;
                }
            }
            //
            //No Viewport type was found with the desired name
            //
            TaskDialog.Show("Warning", "This command uses the Viewport Type " + viewportTypeName + " which doesn't exist in the current model.");
            return;
        }


        public static double[] GetModelDimensionsFrom3D(View3D view3D)
        {
            double[] dimensions = new double[2];
            double minU = 99999;
            double minV = 99999;
            double maxU = -99999;
            double maxV = -99999;

            //
            // Get the orientation of the 3D View
            //
            ViewOrientation3D orientation = view3D.GetOrientation();
            XYZ forward = orientation.ForwardDirection;
            XYZ up = orientation.UpDirection;
            XYZ right = forward.CrossProduct(up);

            XYZ eye = orientation.EyePosition;
            BoundingBoxXYZ bb = view3D.GetSectionBox();

            //
            // List the coordinates of the 3D box
            //
            List<XYZ> coords = new List<XYZ>();
            coords.Add(bb.Min);
            coords.Add(new XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z));
            coords.Add(new XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z));
            coords.Add(new XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z));
            coords.Add(bb.Max);
            coords.Add(new XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z));
            coords.Add(new XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z));
            coords.Add(new XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z));

            List<XYZ> coordsFlat = new List<XYZ>();
            for (int i = 0; i < coords.Count; i++)
            {
                //
                // Project all the coordinates of the section box onto a 2D plane
                //
                XYZ projectedXYZ = MicrodeskHelpers.ProjectPointOnPlane(forward, new XYZ(0, 0, 0), coords[i]);

                //
                // Calculate the 2D UV coorinate
                //
                double u = (projectedXYZ - eye).DotProduct(right);
                double v = (projectedXYZ - eye).DotProduct(up);

                //
                // Update the min and max U's and V's
                //
                if (u < minU)
                    minU = u;

                if (v < minV)
                    minV = v;

                if (u > maxU)
                    maxU = u;

                if (v > maxV)
                    maxV = v;
            }

            //
            // Caluclate the width and height of the view
            //
            dimensions[0] = maxV - minV; //View Height
            dimensions[1] = maxU - minU; //View Width

            return dimensions;
        }


        public static XYZ Calc3DViewInsertion(double[] heightWidth, XYZ bottomLeftJustification, double scale3D)
        {
            double height = heightWidth[0];
            double width = heightWidth[1];

            double u = bottomLeftJustification.X + (width / 2) / scale3D;
            double v = bottomLeftJustification.Y + (height / 2) / scale3D;

            XYZ isoInsertion = new XYZ(u, v, 0);
            
            return isoInsertion;
        }


        public static int GetScale(double modelHeight, double modelWidth, bool fullHeight)
        {
            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("height: {0}", Math.Round(modelHeight, 3));
                writer.WriteLine("width: {0}", Math.Round(modelWidth, 3));
            }

            //
            // Set the max width on the sheet as 8", and max height as 4"
            //
            double paperHeight = 4.0 / 12;
            double paperWidth = 8.0 / 12;

            if (fullHeight) paperHeight = paperHeight * 2.0;

            double bestHeightScale = modelHeight / paperHeight;
            double bestWidthScale = modelWidth / paperWidth;

            double bestScale = bestHeightScale;
            if (bestWidthScale > bestHeightScale)
                bestScale = bestWidthScale;

            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("bestHeightScale: {0}", Math.Round(bestHeightScale, 3));
                writer.WriteLine("bestWidthScale: {0}", Math.Round(bestWidthScale, 3));
            }

            int scale = 96;

            if (bestScale < 1)
                scale = 1;
            else if (bestScale < 2)
                scale = 2;
            else if (bestScale < 4)
                scale = 4;
            else if (bestScale < 8)
                scale = 8;
            else if(bestScale < 12)
                scale = 12;
            else if (bestScale < 16)
                scale = 16;
            else if(bestScale < 24)
                scale = 24;
            else if(bestScale < 32)
                scale = 32;
            else if(bestScale < 48)
                scale = 48;
            else if(bestScale < 72)
                scale = 72;

            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("paperHeight: {0}", Math.Round(paperHeight * scale, 3));
                writer.WriteLine("paperWidth: {0}", Math.Round(paperWidth * scale, 3));
            }

            return scale;
        }


        public static View3D Create3DView(Document doc, ElementId view3DTypeId, BoundingBoxXYZ elementExtents, string spoolNo, ElementId spoolNorthArrowModel)
        {
            // Create a 3D View
            View3D spool3DView = View3D.CreateIsometric(doc, view3DTypeId);


            // Name the View
            spool3DView.get_Parameter(BuiltInParameter.VIEW_NAME).Set(spoolNo + " SouthEast Iso");

            
            // Crop the view
            spool3DView.SetSectionBox(elementExtents);



            // Lock the 3D View
            spool3DView.SaveOrientationAndLock();


            // Set the scale of the 3D View
            double[] projectedHeightWidth = NBKHelpers.GetModelDimensionsFrom3D(spool3DView);


            int scale3D = NBKHelpers.GetScale(projectedHeightWidth[0], projectedHeightWidth[1], true);

            spool3DView.Scale = scale3D;

            // Place the north arrow
            //FamilyInstance northArrow = doc.Create.NewFamilyInstance(
            //    (elementExtents.Min + elementExtents.Max) / 2.0, 
            //    doc.GetElement(spoolNorthArrowModel) as FamilySymbol, 
            //    StructuralType.NonStructural);

            //northArrow.LookupParameter(AppKeys.SpoolParameterName).Set(spoolNo);

            return spool3DView;
        }

        /// <summary>
        /// Create a Revit Section View oritented to the direction of the pipe
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="sectionTypeId"></param>
        /// <param name="elementExtents"></param>
        /// <param name="pipeDirection"></param>
        /// <param name="spoolNo"></param>
        /// <param name="suffix"></param>
        /// <returns></returns>
        public static ViewSection CreateSection(Document doc, ElementId sectionTypeId, BoundingBoxXYZ elementExtents, XYZ pipeDirection, string spoolNo, string suffix)
        {
            // Caluclate the center of the element extents
            double selectionCenterX = (elementExtents.Max.X + elementExtents.Min.X) / 2;
            double selectionCenterY = (elementExtents.Max.Y + elementExtents.Min.Y) / 2;
            double selectionCenterZ = (elementExtents.Max.Z + elementExtents.Min.Z) / 2;

            // Define an 'up' vector
            XYZ up = XYZ.BasisZ;

            // Align the extents
            double angle = pipeDirection.AngleTo(new XYZ(1, 0, 0));
            Transform rot = Transform.CreateRotationAtPoint(XYZ.BasisZ, angle, new XYZ(selectionCenterX, selectionCenterY, 0));
            XYZ alignedExtentsMax = rot.OfPoint(elementExtents.Max);
            XYZ alignedExtentsMin = rot.OfPoint(elementExtents.Min);
            double sectionDepth = Math.Abs(alignedExtentsMax.Z - alignedExtentsMin.Z);
            double sectionWidth = Math.Abs(alignedExtentsMax.X - alignedExtentsMin.X);
            double sectionHeight = Math.Abs(alignedExtentsMax.Y - alignedExtentsMin.Y);

            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine(suffix);
                writer.WriteLine("sectionHeight: {0}", sectionHeight);
                writer.WriteLine("sectionWidth: {0}", sectionWidth);
                writer.WriteLine("sectionDepth: {0}", sectionDepth);
            }


            //Create a bounding box whose origin is at the center of the extents
            //so it goes half the width negative and half the width positive and so on
            XYZ min = new XYZ(-sectionWidth / 2 - .25, 
                -sectionDepth / 2 - .25, 
                -sectionHeight / 2 - .25);
            XYZ max = new XYZ(sectionWidth / 2 + .25,
                sectionDepth / 2 + .25,
                sectionHeight / 2 + .25);



           

            // Create a view direction perpendicular to the pipe
            XYZ viewdir = pipeDirection.CrossProduct(up);


            //t.Origin = new XYZ((elementExtents.Max.X + elementExtents.Min.X) / 2,
            //    (elementExtents.Max.Y + elementExtents.Min.Y) / 2, 
            //    0);
            //t.BasisX = pipeDirection;
            //t.BasisY = up;
            //t.BasisZ = viewdir;

            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("min: {0}, {1}, {2}", Math.Round(min.X, 3), Math.Round(min.Y, 3), Math.Round(min.Z, 3));
                writer.WriteLine("max: {0}, {1}, {2}", Math.Round(max.X, 3), Math.Round(max.Y, 3), Math.Round(max.Z, 3));
            }

            BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
            sectionBox.Min = min;
            sectionBox.Max = max;

            //Create a transform and apply it to the section box
            Transform t = Transform.Identity;
            t.Origin = new XYZ(selectionCenterX, selectionCenterY, selectionCenterZ);
            t.BasisX = pipeDirection;
            t.BasisY = up;
            t.BasisZ = viewdir;
            sectionBox.Transform = t;

            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("sectionBox.Min: {0}, {1}, {2}", Math.Round(sectionBox.Min.X, 3), Math.Round(sectionBox.Min.Y, 3), Math.Round(sectionBox.Min.Z, 3));
                writer.WriteLine("sectionBox.max: {0}, {1}, {2}", Math.Round(sectionBox.Max.X, 3), Math.Round(sectionBox.Max.Y, 3), Math.Round(sectionBox.Max.Z, 3));
                writer.WriteLine("t.Origin: {0}, {1}, {2}", Math.Round(t.Origin.X, 3), Math.Round(t.Origin.Y, 3), Math.Round(t.Origin.Z, 3));
                writer.WriteLine("t.BasisX: {0}, {1}, {2}", Math.Round(t.BasisX.X, 3), Math.Round(t.BasisX.Y, 3), Math.Round(t.BasisX.Z, 3));
                writer.WriteLine("t.BasisY: {0}, {1}, {2}", Math.Round(t.BasisY.X, 3), Math.Round(t.BasisY.Y, 3), Math.Round(t.BasisY.Z, 3));
                writer.WriteLine("t.BasisZ: {0}, {1}, {2}", Math.Round(t.BasisZ.X, 3), Math.Round(t.BasisZ.Y, 3), Math.Round(t.BasisZ.Z, 3));
            }


            ViewSection spoolSectionX = ViewSection.CreateSection(doc, sectionTypeId, sectionBox);



            // Crop the View
            spoolSectionX.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE).Set(1);
            spoolSectionX.get_Parameter(BuiltInParameter.VIEWER_CROP_REGION_VISIBLE).Set(0);



            // Set the scale
            int scaleSection = NBKHelpers.GetScale(sectionHeight, sectionWidth, false);
            spoolSectionX.Scale = scaleSection;

            // Name the View
            spoolSectionX.get_Parameter(BuiltInParameter.VIEW_NAME).Set("Spool " + spoolNo + " " + suffix);

            return spoolSectionX;
        }


        public static ViewSection CreateSection_Backup(Document doc, ElementId sectionTypeId, BoundingBoxXYZ elementExtents, XYZ pipeDirection, string spoolNo, string suffix)
        {


            XYZ min = new XYZ(-((elementExtents.Max.X - elementExtents.Min.X) / 2) - 1, elementExtents.Min.Z, -(elementExtents.Max.Y - elementExtents.Min.Y));
            XYZ max = new XYZ(((elementExtents.Max.X - elementExtents.Min.X) / 2) + 1, elementExtents.Max.Z, (elementExtents.Max.Y - elementExtents.Min.Y));



            XYZ up = XYZ.BasisZ;

            // Create a view direction perpendicular to the pipe
            XYZ viewdir = pipeDirection.CrossProduct(up);


            //t.Origin = new XYZ((elementExtents.Max.X + elementExtents.Min.X) / 2,
            //    (elementExtents.Max.Y + elementExtents.Min.Y) / 2, 
            //    0);
            //t.BasisX = pipeDirection;
            //t.BasisY = up;
            //t.BasisZ = viewdir;


            BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
            sectionBox.Min = min;
            sectionBox.Max = max;

            //Create a transform and apply it to the section box
            double selectionCenterX = (elementExtents.Max.X + elementExtents.Min.X) / 2;
            double selectionCenterY = (elementExtents.Max.Y + elementExtents.Min.Y) / 2;
            Transform t = Transform.Identity;
            t.Origin = new XYZ(selectionCenterX, selectionCenterY, 0);
            t.BasisX = pipeDirection;
            t.BasisY = up;
            t.BasisZ = viewdir;
            sectionBox.Transform = t;

            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("t.Origin: {0}, {1}, {2}", Math.Round(t.Origin.X, 3), Math.Round(t.Origin.Y, 3), Math.Round(t.Origin.Z, 3));
                writer.WriteLine("t.BasisX: {0}, {1}, {2}", Math.Round(t.BasisX.X, 3), Math.Round(t.BasisX.Y, 3), Math.Round(t.BasisX.Z, 3));
                writer.WriteLine("t.BasisY: {0}, {1}, {2}", Math.Round(t.BasisY.X, 3), Math.Round(t.BasisY.Y, 3), Math.Round(t.BasisY.Z, 3));
                writer.WriteLine("t.BasisZ: {0}, {1}, {2}", Math.Round(t.BasisZ.X, 3), Math.Round(t.BasisZ.Y, 3), Math.Round(t.BasisZ.Z, 3));
            }


            ViewSection spoolSectionX = ViewSection.CreateSection(doc, sectionTypeId, sectionBox);



            // Crop the View
            spoolSectionX.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE).Set(1);
            spoolSectionX.get_Parameter(BuiltInParameter.VIEWER_CROP_REGION_VISIBLE).Set(0);



            // Set the scale
            int scaleSection = NBKHelpers.GetScale(elementExtents.Max.Z - elementExtents.Min.Z, elementExtents.Max.X - elementExtents.Min.X, false);
            spoolSectionX.Scale = scaleSection;

            // Name the View
            spoolSectionX.get_Parameter(BuiltInParameter.VIEW_NAME).Set("Spool " + spoolNo + " " + suffix);

            return spoolSectionX;
        }


        public static ViewPlan CreatePlan(Document doc, 
            ElementId planTypeId, 
            BoundingBoxXYZ elementExtents, 
            string spoolNo, 
            XYZ pipeDirection, 
            ElementId northArrowSymbolId)
        {
            // Find the best level to base the plan on
            ElementId levelId = MicrodeskHelpers.GetLevel(doc, elementExtents.Min.Z);



            ViewPlan spoolPlan = ViewPlan.Create(doc, planTypeId, levelId);

            // Name View
            spoolPlan.get_Parameter(BuiltInParameter.VIEW_NAME).Set("Spool " + spoolNo + " Plan");

            XYZ min = elementExtents.Min;
            XYZ max = elementExtents.Max;
            BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
            sectionBox.Min = min;
            sectionBox.Max = max;

#region log
            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("sectionBox.Min: {0}, {1}, {2}", Math.Round(min.X, 2), Math.Round(min.Y, 2), Math.Round(min.Z, 2));
                writer.WriteLine("sectionBox.Min: {0}, {1}, {2}", Math.Round(max.X, 2), Math.Round(max.Y, 2), Math.Round(max.Z, 2));
            }
#endregion log

            // Rotate the View
            if (pipeDirection != null)
            {
#region log
                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                {
                    writer.WriteLine("Plan pipeDirection: {0}, {1}, {2}", Math.Round(pipeDirection.X, 2), Math.Round(pipeDirection.Y, 2), Math.Round(pipeDirection.Z, 2));
                }
#endregion log

                XYZ up = XYZ.BasisZ;


                //Create a transform and apply it to the section box
                double selectionCenterX = (elementExtents.Max.X + elementExtents.Min.X) / 2;
                double selectionCenterY = (elementExtents.Max.Y + elementExtents.Min.Y) / 2;
                double selectionCenterZ = (elementExtents.Max.Z + elementExtents.Min.Z) / 2;
                Transform t = Transform.Identity;
                t.Origin = new XYZ(selectionCenterX, selectionCenterY, selectionCenterZ);
                t.BasisX = pipeDirection;
                t.BasisY = -1.0 * pipeDirection.CrossProduct(up); 
                t.BasisZ = up;
                
                sectionBox.Transform = t;
            }

            // Crop the View
            spoolPlan.CropBoxActive = true;
            spoolPlan.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE).Set(1);
            spoolPlan.get_Parameter(BuiltInParameter.VIEWER_CROP_REGION_VISIBLE).Set(0);
            spoolPlan.CropBox = sectionBox;


            // Set the View Range
            PlanViewRange viewRange = spoolPlan.GetViewRange();
            Level level = doc.GetElement(levelId) as Level;
            double levelElevation = (level).Elevation;
            double bottomOffset = elementExtents.Min.Z - levelElevation - 1;
            double topOffset = elementExtents.Max.Z - levelElevation + 1;
            viewRange.SetOffset(PlanViewPlane.BottomClipPlane, bottomOffset);
            viewRange.SetOffset(PlanViewPlane.ViewDepthPlane, bottomOffset);
            viewRange.SetOffset(PlanViewPlane.TopClipPlane, topOffset);
            viewRange.SetOffset(PlanViewPlane.CutPlane, topOffset);
            spoolPlan.SetViewRange(viewRange);



            // Set the scale
            int scalePlan = NBKHelpers.GetScale(elementExtents.Max.Y - elementExtents.Min.Y, elementExtents.Max.X - elementExtents.Min.X, false);
            spoolPlan.Scale = scalePlan;

            // Place the north arrow
            doc.Create.NewFamilyInstance(sectionBox.Min, doc.GetElement(northArrowSymbolId) as FamilySymbol, spoolPlan);

            return spoolPlan;
        }


        public static void RotateBoundingBox(Document doc, ViewPlan plan, double angle)
        {
#region log
            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("RotateBoundingBox: {0}", Math.Round(angle * 180 / Math.PI, 2));
            }
#endregion log


            BoundingBoxXYZ box = plan.CropBox;

            XYZ center = 0.5 * (box.Max + box.Min);

            Line axis = Line.CreateBound(
              center, center + XYZ.BasisZ);

            ElementId cropBoxId = GetCropBoxFor(plan);

            ElementTransformUtils.RotateElement(doc,
              cropBoxId, axis, angle);
        }


        public static ElementId GetCropBoxFor(ViewPlan plan)
        {
            ParameterValueProvider provider
              = new ParameterValueProvider(new ElementId(
                (int)BuiltInParameter.ID_PARAM));

            FilterElementIdRule rule
              = new FilterElementIdRule(provider,
                new FilterNumericEquals(), plan.Id);

            ElementParameterFilter filter
              = new ElementParameterFilter(rule);

            return new FilteredElementCollector(plan.Document)
              .WherePasses(filter)
              .ToElementIds()
              .Where<ElementId>(a => a.IntegerValue
               != plan.Id.IntegerValue)
              .FirstOrDefault<ElementId>();
        }


        public static void TagAllInView(Document doc, 
            ElementId taggIdPipe, 
            ElementId tagIdFitting, 
            List<RevitElement> selectedElements, 
            Autodesk.Revit.DB.View taggingView, 
            View3D spool3DView, 
            bool is3D)
        {
            foreach (RevitElement r in selectedElements)
            {

                BoundingBoxXYZ bb2 = doc.GetElement(r.ElementId).get_BoundingBox(spool3DView);

                XYZ tagLocation = r.Center;

                if (is3D)
                    tagLocation = GetTagPosition(spool3DView, r.Center);


                IndependentTag newTag = IndependentTag.Create(doc, tagIdFitting, taggingView.Id, r.Reference, true, TagOrientation.Horizontal, r.Center);
                if(r.PartType == "Pipe")
                {
                    newTag.ChangeTypeId(taggIdPipe);
                }



                //IndependentTag newTag = IndependentTag.Create(doc, taggingView.Id, r.Reference, true, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, r.Center);

                newTag.LeaderEndCondition = LeaderEndCondition.Free;
                newTag.LeaderEnd = r.Center;

                //IndependentTag newTag = IndependentTag.Create(m_document, tagId, spool3DView.Id, r.Reference, false, TagOrientation.Horizontal, r.Center);
#region log
                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                {
                    writer.WriteLine("Original BB: {0}, {1}", Math.Round(r.Min.X, 2), Math.Round(r.Min.Y, 2));
                    writer.WriteLine("New View BB: {0}, {1}", Math.Round(bb2.Min.X, 2), Math.Round(bb2.Min.Y, 2));
                    writer.WriteLine("Tag Created: {0}", newTag.Id);
                }
#endregion log
            }
        }


        public static void TagAllIn3DView(Document doc, List<RevitElement> selectedElements, Autodesk.Revit.DB.View taggingView, View3D spool3DView, bool is3D)
        {
            foreach (RevitElement r in selectedElements)
            {

                BoundingBoxXYZ bb2 = doc.GetElement(r.ElementId).get_BoundingBox(taggingView);

                XYZ center = new XYZ((bb2.Min.X + bb2.Max.X) / 2, (bb2.Min.Y + bb2.Max.Y) / 2, (bb2.Min.Z + bb2.Max.Z) / 2);
                XYZ leaderLocation = r.Center;

                if (is3D)
                    leaderLocation = GetTagPosition(spool3DView, r.Center);




                IndependentTag newTag = IndependentTag.Create(doc, 
                    taggingView.Id, 
                    r.Reference, 
                    true, 
                    TagMode.TM_ADDBY_CATEGORY, 
                    TagOrientation.Horizontal, 
                    leaderLocation + new XYZ(.1, .1, .1));

                newTag.HasLeader = true;
                newTag.LeaderEndCondition = LeaderEndCondition.Free;
                newTag.LeaderEnd = leaderLocation;
                

                //IndependentTag newTag = IndependentTag.Create(m_document, tagId, spool3DView.Id, r.Reference, false, TagOrientation.Horizontal, r.Center);
#region log
                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                {
                    writer.WriteLine("Original BB: {0}, {1}", Math.Round(r.Min.X, 2), Math.Round(r.Min.Y, 2));
                    writer.WriteLine("New View BB: {0}, {1}", Math.Round(bb2.Min.X, 2), Math.Round(bb2.Min.Y, 2));
                    writer.WriteLine("Original Center: {0}, {1}, {2}", Math.Round(r.Center.X, 2), Math.Round(r.Center.Y, 2), Math.Round(r.Center.Z, 2));
                    writer.WriteLine("New Center: {0}, {1}, {2}", Math.Round(center.X, 2), Math.Round(center.Y, 2), Math.Round(center.Z, 2));
                    writer.WriteLine("Tag Created: {0}", newTag.Id);
                }
#endregion log
            }
        }


        public static XYZ GetTagPosition(View3D spool3DView, XYZ center)
        {
            //
            // get the orientation of the 3d view
            //
            ViewOrientation3D orientation = spool3DView.GetOrientation();
            XYZ forward = orientation.ForwardDirection;
            XYZ up = orientation.UpDirection;
            XYZ right = forward.CrossProduct(up);

            XYZ eye = orientation.EyePosition;

            XYZ projectedXYZ = MicrodeskHelpers.ProjectPointOnPlane(forward, eye, center);

            return projectedXYZ;
        }

        public static void TagAdjacentSpools(Document doc, List<RevitElement> selectedElements, string spoolNo)
        {
            foreach(RevitElement re in selectedElements)
            {
                Connector[] connectors = MicrodeskHelpers.ConnectorArray(re.Element);

                int i = 0;
                foreach(Connector c in connectors)
                {
                    using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                    {
                        writer.WriteLine("Selected Element {0} connector {1}", re.Element.Id.ToString(), i);
                    }

                    // skip unconnected connectors
                    //if (!c.IsConnected) continue;

                    ConnectorSet connectedConnectors = c.AllRefs;

                    foreach(Connector connectedConnector in connectedConnectors)
                    {
                        using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                        {
                            writer.WriteLine("Selected Element {0} connector {1} connected connector {2}", re.Element.Id.ToString(), i, connectedConnector.Owner.Id.ToString());
                        }

                        if (connectedConnector.Owner.Id.ToString() == re.ElementId.ToString()) continue;

                        bool outsideOfSelectedElements = true;

                        foreach (RevitElement revElem in selectedElements)
                        {
                            if(connectedConnector.Owner.Id.ToString() == revElem.ElementId.ToString())
                            {
                                outsideOfSelectedElements = false;
                                break;
                            }
                        }

                        using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                        {
                            writer.WriteLine("outsideOfSelectedElements {0}", outsideOfSelectedElements.ToString());
                        }

                        if (outsideOfSelectedElements)
                        {
                            Parameter spoolAdj1Param = connectedConnector.Owner.LookupParameter(AppKeys.SpoolAdjacent1ParameterName);
                            Parameter spoolAdj2Param = connectedConnector.Owner.LookupParameter(AppKeys.SpoolAdjacent2ParameterName);
                            string spoolAdjString = spoolNo;

                            if (spoolAdj1Param.AsString() != null &&
                                spoolAdj1Param.AsString() != "")
                            {
                                spoolAdj2Param.Set(spoolNo);
                            }
                            else
                            {
                                spoolAdj1Param.Set(spoolNo);
                            }
                        }
                    }
                    i++;
                }
            }
        }

        public static int WriteColumnHeadersPipe(Excel.Worksheet xlWorkSheet, int activeRow)
        {
            xlWorkSheet.Cells[activeRow, 1] = "Long Description";
            xlWorkSheet.Cells[activeRow, 2] = "Size";
            xlWorkSheet.Cells[activeRow, 3] = "Count";
            xlWorkSheet.Cells[activeRow, 4] = "Length";
            xlWorkSheet.Cells[activeRow, 5] = "Pipe Length";
            xlWorkSheet.Cells[activeRow, 6] = "Manufacturer Name";
            xlWorkSheet.Cells[activeRow, 7] = "Product Line";
            xlWorkSheet.Cells[activeRow, 8] = "Short Description";
            xlWorkSheet.Cells[activeRow, 9] = "Product";
            xlWorkSheet.Cells[activeRow, 10] = "Material";

            Excel.Range range1 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 1], xlWorkSheet.Cells[activeRow + 1, 1]];
            range1.Cells.ColumnWidth = 39.43;
            Excel.Range range2 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 2], xlWorkSheet.Cells[activeRow + 1, 2]];
            range2.Cells.ColumnWidth = 7.86;
            Excel.Range range3 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 3], xlWorkSheet.Cells[activeRow + 1, 3]];
            range3.Cells.ColumnWidth = 9.29;
            Excel.Range range4 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 4], xlWorkSheet.Cells[activeRow + 1, 4]];
            range4.Cells.ColumnWidth = 9.29;
            Excel.Range range5 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 5], xlWorkSheet.Cells[activeRow + 1, 5]];
            range5.Cells.ColumnWidth = 12.14;
            Excel.Range range6 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 6], xlWorkSheet.Cells[activeRow + 1, 6]];
            range6.Cells.ColumnWidth = 17.86;
            Excel.Range range7 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 7], xlWorkSheet.Cells[activeRow + 1, 7]];
            range7.Cells.ColumnWidth = 17.43;
            Excel.Range range8 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 8], xlWorkSheet.Cells[activeRow + 1, 8]];
            range8.Cells.ColumnWidth = 15;
            Excel.Range range9 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 9], xlWorkSheet.Cells[activeRow + 1, 9]];
            range9.Cells.ColumnWidth = 9.29;
            Excel.Range range10 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 10], xlWorkSheet.Cells[activeRow + 1, 10]];
            range10.Cells.ColumnWidth = 12;

            Excel.Range range0 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 1], xlWorkSheet.Cells[activeRow, 10]];
            range0.Cells.RowHeight = 17.25;
            range0.Cells.Font.Bold = true;
            range0.Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(169, 169, 169));

            activeRow++;
            return activeRow;
        }

        public static int WriteColumnHeadersFittings(Excel.Worksheet xlWorkSheet, int activeRow)
        {
            xlWorkSheet.Cells[activeRow, 1] = "Long Description";
            xlWorkSheet.Cells[activeRow, 2] = "Size";
            xlWorkSheet.Cells[activeRow, 3] = "Count";
            xlWorkSheet.Cells[activeRow, 4] = "Manufacturer Name";
            xlWorkSheet.Cells[activeRow, 5] = "Product Line";
            xlWorkSheet.Cells[activeRow, 6] = "Short Description";
            xlWorkSheet.Cells[activeRow, 7] = "Product";
            xlWorkSheet.Cells[activeRow, 8] = "Material";

            Excel.Range range1 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 1], xlWorkSheet.Cells[activeRow + 1, 1]];
            range1.Cells.ColumnWidth = 39.43;
            Excel.Range range2 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 2], xlWorkSheet.Cells[activeRow + 1, 2]];
            range2.Cells.ColumnWidth = 7.86;
            Excel.Range range3 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 3], xlWorkSheet.Cells[activeRow + 1, 3]];
            range3.Cells.ColumnWidth = 9.29;
            Excel.Range range6 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 4], xlWorkSheet.Cells[activeRow + 1, 4]];
            range6.Cells.ColumnWidth = 17.86;
            Excel.Range range7 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 5], xlWorkSheet.Cells[activeRow + 1, 5]];
            range7.Cells.ColumnWidth = 17.43;
            Excel.Range range8 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 6], xlWorkSheet.Cells[activeRow + 1, 6]];
            range8.Cells.ColumnWidth = 15;
            Excel.Range range9 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 7], xlWorkSheet.Cells[activeRow + 1, 7]];
            range9.Cells.ColumnWidth = 9.29;
            Excel.Range range10 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 8], xlWorkSheet.Cells[activeRow + 1, 8]];
            range10.Cells.ColumnWidth = 12;

            Excel.Range range0 = xlWorkSheet.Range[xlWorkSheet.Cells[activeRow, 1], xlWorkSheet.Cells[activeRow, 8]];
            range0.Cells.RowHeight = 17.25;
            range0.Cells.Font.Bold = true;
            range0.Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(169, 169, 169));

            activeRow++;
            return activeRow;
        }

        public static void AddBorders(Excel.Worksheet xlWorkSheet)
        {
            Excel.Range usedCells = xlWorkSheet.UsedRange;

            BorderAround(usedCells, 1);

        }

        private static void BorderAround(Excel.Range range, int colour)
        {
            Excel.Borders borders = range.Borders;
            borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders.Color = colour;
            borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            borders = null;
        }
    }
}
