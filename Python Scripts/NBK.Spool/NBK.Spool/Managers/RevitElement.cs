using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.Revit.DB;
using Element = Autodesk.Revit.DB.Element;

namespace NBK.Spool
{
    public class RevitElement
    {
        public string Name { get; set; }
        public string FamilyName { get; set; }
        public string FamilyTypeName { get; set; }
        public string Size { get; set; }
        public string FamilyandSize { get; set; }
        public string FamilySizeandLength { get; set; }
        public string CategoryName { get; set; }
        public string PartType { get; set; }
        public string SpoolNumber { get; set; }
        public string PipeLength { get; set; }
        public string ProductLine { get; set; }
        public string ShortDescription { get; set; }
        public string Manufacturer { get; set; }
        public string Product { get; set; }
        public string Material { get; set; }
        public string Connector0Name { get; set; }
        public string Connector1Name { get; set; }
        public Element Element { get; set; }
        public ElementId ElementId { get; set; }
        public ElementId LinkInstanceId { get; set; }
        public Document LinkDocument { get; set; }
        public Reference Reference { get; set; }
        public List<XYZ> PartDirections { get; set; }
        public XYZ PartDirection { get; set; }

        public Connector[] Connectors { get; set; }
        public BoundingBoxXYZ BoundingBox { get; set; }
        public XYZ Min { get; set; }
        public XYZ Max { get; set; }
        public XYZ Center { get; set; }
        public double Length { get; set; }
        public double Height { get; set; }
        public double Width { get; set; }
        public double Elevation { get; set; }
        public double ElevationRounded { get; set; }
        public string LevelDisplay { get; set; }
        public string FloorCode { get; set; }

        public RevitElement()
        {
        }

        public RevitElement(ElementId id, string name, Document linkDocument)
        {
            this.LinkInstanceId = id;
            this.Name = name;
            this.LinkDocument = linkDocument;
        }

        public RevitElement(Reference reference, 
            Element element, 
            string familyName, 
            string familyTypeName, 
            string partType,
            string size, 
            string length, 
            BoundingBoxXYZ boundingBox,
            List<XYZ> directions,
            string connector0Name,
            string connector1Name)
        {
            this.Reference = reference;
            this.ElementId = element.Id;
            this.Element = element;
            this.FamilyName = familyName;
            this.FamilyTypeName = familyTypeName;
            this.Size = size;
            this.FamilyandSize = familyName + "|" + size;
            this.FamilySizeandLength = familyName + "|" + size + "|" + length;
            this.PartType = partType;
            this.BoundingBox = boundingBox;
            this.Min = boundingBox.Min;
            this.Max = boundingBox.Max;
            this.Center = (boundingBox.Min + boundingBox.Max) / 2;
            this.PartDirections = directions;
            this.Connector0Name = connector0Name;
            this.Connector1Name = connector1Name;
        }

        public RevitElement(Element element,
            string familyName,
            string familyTypeName,
            string partType,
            string size,
            double length,
            string pipeLength,
            string spoolNo,
            string manufacturer,
            string productLine,
            string shortDescription,
            string product,
            string material)
        {
            this.ElementId = element.Id;
            this.Element = element;
            this.FamilyName = familyName;
            this.FamilyTypeName = familyTypeName;
            this.Size = size;
            this.Length = length;
            this.PipeLength = pipeLength;
            this.FamilyandSize = familyName + "|" + size;
            this.FamilySizeandLength = familyName + "|" + size + "|" + length;
            this.PartType = partType;
            this.SpoolNumber = spoolNo;
            this.Manufacturer = manufacturer;
            this.ProductLine = productLine;
            this.ShortDescription = shortDescription;
            this.Product = product;
            this.Material = material;
        }

        public RevitElement(string name, double elevation)
        {
            this.Name = name;
            this.Elevation = elevation;
            this.ElevationRounded = Math.Round(elevation, 2);
            this.LevelDisplay = this.ElevationRounded.ToString() + " ft,    " + name;
        }

        public RevitElement(ElementId id, string name, string catName, Document linkDocument, Connector[] connectors, XYZ min, XYZ max, double height, double width)
        {
            this.ElementId = id;
            this.Name = name;
            this.CategoryName = catName;
            this.LinkDocument = linkDocument;
            this.Connectors = connectors;
            this.Min = min;
            this.Max = max;
            this.Height = height;
            this.Width = width;
        }

        public RevitElement(string name, XYZ direction)
        {
            this.Name = name;
            this.PartDirection = direction;
        }
    }
}
