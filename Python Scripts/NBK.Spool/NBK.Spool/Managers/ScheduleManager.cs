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
    public class ScheduleManager
    {
        protected Autodesk.Revit.ApplicationServices.Application m_application;
        protected Document m_document;
        protected Element m_exampleFamily;
        protected Parameter m_parameter;
        protected BuiltInCategory m_builtinCategory;
        protected ElementId m_categoryId;
        protected string m_spoolNo;
        protected string m_partType;
        protected SchedulableField m_lengthSchedField;
        protected SchedulableField m_sizeSchedField;
        protected SchedulableField m_connector0SchedField;
        protected SchedulableField m_connector1SchedField;
        protected ElementId m_viewTypeId;
        protected ElementId m_viewTemplateId;
        public ViewSchedule ViewSchedule;
        

        public ScheduleManager(Autodesk.Revit.ApplicationServices.Application app, 
            Document doc, 
            BuiltInCategory builtInCategory, 
            Element element,
            string spoolNo,
            string partType,
            ElementId scheduleTypeId,
            ElementId viewTemplateId)
        {
            m_application = app;
            m_document = doc;
            m_exampleFamily = element;
            m_builtinCategory = builtInCategory;
            m_categoryId = new ElementId(builtInCategory);
            m_spoolNo = spoolNo;
            m_partType = partType;
            m_viewTypeId = scheduleTypeId;
            m_viewTemplateId = viewTemplateId;

            CreatePipeSchedule();
        }

        public void CreatePipeSchedule()
        {
            // Create a schedule
            this.ViewSchedule = ViewSchedule.CreateSchedule(m_document, m_categoryId);

            #region Log
            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("ViewSchedule: {0}", this.ViewSchedule.Id);
            }
            #endregion Log


            List<ScheduleFieldInfo> scheduleFields = ListScheduleFields();

            #region Log
            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
            {
                writer.WriteLine("scheduleFields: {0}", scheduleFields.Count);
            }
            #endregion Log

            AddFieldToSchedule(this.ViewSchedule, 
                scheduleFields,
                AppKeys.SpoolItemParameterName);

            // We have to add 'Size' and 'Length' separately to the schedule because some fabrication parts don't contain it.
            if (this.m_sizeSchedField != null)
            {
                ScheduleField sizeColumn = this.ViewSchedule.Definition.AddField(this.m_sizeSchedField);
                sizeColumn.SheetColumnWidth = 1.0 / 12.0;
            }


            if (this.m_lengthSchedField != null  && this.m_partType == "Pipe")
            {
                ScheduleField lengthColumn = this.ViewSchedule.Definition.AddField(this.m_lengthSchedField);
                lengthColumn.SheetColumnWidth = 1.0 / 12.0;
            }

            if (this.m_connector0SchedField != null && this.m_partType == "Pipe")
            {
                ScheduleField conn0Column = this.ViewSchedule.Definition.AddField(this.m_connector0SchedField);
                conn0Column.SheetColumnWidth = 1.0 / 12.0;
                conn0Column.ColumnHeading = "Connector 1";
            }

            if (this.m_connector1SchedField != null && this.m_partType == "Pipe")
            {
                ScheduleField conn1Column = this.ViewSchedule.Definition.AddField(this.m_connector1SchedField);
                conn1Column.SheetColumnWidth = 1.0 / 12.0;
                conn1Column.ColumnHeading = "Connector 2";
            }




            // Add count to schedule
            ScheduleField countColumn = this.ViewSchedule.Definition.AddField(ScheduleFieldType.Count);
            countColumn.ColumnHeading = "QTY";
            countColumn.SheetColumnWidth = 0.5 / 12.0;

            // Name the Schedule
            if (m_partType == null)
            {
                this.ViewSchedule.get_Parameter(BuiltInParameter.VIEW_NAME).Set("SPOOL TOTAL WEIGHT " + m_spoolNo);
            }
            else
            {
                this.ViewSchedule.get_Parameter(BuiltInParameter.VIEW_NAME).Set(m_partType.ToUpper() + " SPOOL NO. " + m_spoolNo);
            }

            // Hide the first two columns
            //this.ViewSchedule.Definition.GetField(0).IsHidden = true;
            //this.ViewSchedule.Definition.GetField(1).IsHidden = true;


            // Apply the View Template
            this.ViewSchedule.ViewTemplateId = m_viewTemplateId;

            // Change the View Type
            this.ViewSchedule.ChangeTypeId(m_viewTypeId);
        }

        private Parameter GetParameter(string parameterName)
        {
            Parameter p = m_exampleFamily.LookupParameter(parameterName);

            if(null==p)
            {
                TaskDialog.Show("Warning", "The Bill of Materials requires a Parameter named " + parameterName);
            }

            return p;
        }


        private List<ScheduleFieldInfo> ListScheduleFields()
        {
            List<ScheduleFieldInfo> fieldList = new List<ScheduleFieldInfo>();

            fieldList.Add(new ScheduleFieldInfo("Spool", AppKeys.SpoolParameterName, 0.5));
            fieldList.Add(new ScheduleFieldInfo("Type", AppKeys.SpoolTypeParameterName, 0.5));
            fieldList.Add(new ScheduleFieldInfo("Piece Number", AppKeys.SpoolItemParameterName, 0.75));
            fieldList.Add(new ScheduleFieldInfo("Description", "Family", 4.0));
            //fieldList.Add(new ScheduleFieldInfo("Length", "Length", 1.5));
            //fieldList.Add(new ScheduleFieldInfo("QTY", "Count"));
            //fieldList.Add(new ScheduleFieldInfo("QTY", "Spool Quantity", GetParameter("Spool Quantity")));
            //fieldList.Add(new ScheduleFieldInfo("Length (ft)", "Spool Length", GetParameter("Spool Length")));
            //fieldList.Add(new ScheduleFieldInfo("Item", "Spool Description", GetParameter("Spool Description")));

            foreach (ScheduleFieldInfo sFI in fieldList)
            {
                sFI.Parameter = m_exampleFamily.LookupParameter(sFI.ParameterName);
            }

            return fieldList;
        }


        private static List<Parameter> EnterParameterValue(string[] paramNameArray, 
            string[] paramValueArray, 
            List<Parameter> updatedParameters, 
            FamilyInstance fInst, 
            FamilySymbol fSym, 
            int j, 
            string log)
        {
            Parameter currentParam = fInst.LookupParameter(paramNameArray[j]);
            if (currentParam == null)
            {
                currentParam = fSym.LookupParameter(paramNameArray[j]);

            }



            //Isolate string parameters
            else if (currentParam.StorageType == StorageType.String)
            {
                updatedParameters.Add(currentParam);

                currentParam.Set(paramValueArray[j]);
            }


            //Isolate double parameters
            else if (currentParam.StorageType == StorageType.Double)
            {
                updatedParameters.Add(currentParam);

                if (currentParam.DisplayUnitType == DisplayUnitType.DUT_FAHRENHEIT)
                {
                    double temperature = 273.15 + (0 - 32) * 5 / 9;
                    try
                    {
                        temperature = 273.15 + (Convert.ToDouble(paramValueArray[j]) - 32) * 5 / 9;
                    }
                    catch { }
                    currentParam.Set(temperature);
                }
                else if (currentParam.DisplayUnitType == DisplayUnitType.DUT_CUBIC_FEET_PER_MINUTE)
                {
                    double convertedCFM = Convert.ToDouble(paramValueArray[j]) / 60;
                    currentParam.Set(convertedCFM);
                }
                else if (currentParam.DisplayUnitType == DisplayUnitType.DUT_GALLONS_US_PER_MINUTE)
                {
                    double convertedGPM = Convert.ToDouble(paramValueArray[j]) * 0.00222800938;
                    currentParam.Set(convertedGPM);
                }
                else if (currentParam.DisplayUnitType == DisplayUnitType.DUT_INCHES_OF_WATER)
                {
                    double convertedPress = Convert.ToDouble(paramValueArray[j]) / .013171;
                    currentParam.Set(convertedPress);
                }
                else if (currentParam.DisplayUnitType == DisplayUnitType.DUT_VOLTS)
                {
                    double convertedVolts = Convert.ToDouble(paramValueArray[j]) * 10.76233;
                    currentParam.Set(convertedVolts);
                }
                else if (currentParam.DisplayUnitType == DisplayUnitType.DUT_WATTS)
                {
                    double convertedWatts = Convert.ToDouble(paramValueArray[j]) / 10.76239;
                    currentParam.Set(convertedWatts);
                }
                else
                {
                    currentParam.Set(Convert.ToDouble(paramValueArray[j]));
                }
            }
            //Isolate integer parameters
            else if (currentParam.StorageType == StorageType.Integer)
            {
                updatedParameters.Add(currentParam);

                currentParam.Set(Convert.ToInt16(paramValueArray[j]));
            }



            return updatedParameters;
        }



        //
        //Add fields to view schedule.
        //
        public void AddFieldToSchedule(ViewSchedule vs, 
            List<ScheduleFieldInfo> parametersForSchedule,
            string sortParameterName)
        {
            IList<SchedulableField> schedulableFields = null;


            //
            //Get all schedulable fields from view schedule definition.
            //
            schedulableFields = vs.Definition.GetSchedulableFields();


            for (int i = 0; i < parametersForSchedule.Count; i++)
            {
                Parameter currentParam = m_exampleFamily.LookupParameter(parametersForSchedule[i].ParameterName);
                //Parameter currentParam = parametersForSchedule[i].Parameter;
                string columnHeader = parametersForSchedule[i].ColumnHeader;


                //
                // Loop through each of the fields available to schedule
                //
                foreach (SchedulableField sf in schedulableFields)
                {
                    if (sf.GetName(m_document) == "Length")
                    {
                        this.m_lengthSchedField = sf;
                        continue;
                    }

                    if (sf.GetName(m_document) == "Size")
                    {
                        this.m_sizeSchedField = sf;
                        continue;
                    }

                    if (sf.GetName(m_document) == AppKeys.SpoolConn0ParameterName)
                    {
                        this.m_connector0SchedField = sf;
                        continue;
                    }

                    if (sf.GetName(m_document) == AppKeys.SpoolConn1ParameterName)
                    {
                        this.m_connector1SchedField = sf;
                        continue;
                    }

                    if (currentParam.Id == sf.ParameterId)
                    {
                        string schedulableFieldName = sf.GetName(m_document);

                        //
                        //The same parameter id could belong to rooms and spaces, which are also schedulable
                        //though we don't want them
                        //
                        if (schedulableFieldName.Contains("Room:") || schedulableFieldName.Contains("Space:"))
                        {
                            continue;
                        }

                        try
                        {
                            //Add the schedulable field to the schedule
                            ScheduleField field = vs.Definition.AddField(sf);

                            //Name the column Header
                            field.ColumnHeading = columnHeader;

                            // Filter the schedule

                            // Filter first by the spool number
                            if (schedulableFieldName == AppKeys.SpoolParameterName)
                            {
                                ScheduleFilter schedFilter = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, m_spoolNo);
                                vs.Definition.AddFilter(schedFilter);
                            }

                            // Filter second by the part type
                            // 'Pipe' or 'Pipe Fitting'.
                            // Do not filter if part type was entered as 'null'
                            // like in a weight schedule
                            if (m_partType != null && 
                                schedulableFieldName == AppKeys.SpoolTypeParameterName)
                            {
                                ScheduleFilter schedFilter = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, m_partType);
                                vs.Definition.AddFilter(schedFilter);
                            }

                            // Group and sort the view schedule by column A, normally the tag
                            if (schedulableFieldName == sortParameterName)
                            {
                                ScheduleSortGroupField sortGroupField = new ScheduleSortGroupField(field.FieldId);
                                sortGroupField.ShowBlankLine = false;
                                vs.Definition.AddSortGroupField(sortGroupField);

                                // Uncheck "Itemize Every Instance"
                                // in the case of Pipe Fitting and Weight schedules
                                // Pipe schedules should be itemized, though
                                if(AppKeys.SpoolPipe != m_partType)
                                    vs.Definition.IsItemized = false;
                            }

                            field.SheetColumnWidth = parametersForSchedule[i].ColumWidthFeet;
                        }
                        catch
                        {
                            #region Log
                            using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                            {
                                writer.WriteLine("Field Previously Added");
                            }
                            #endregion Log
                        }
                    }
                }
            }
        }

    }
}
