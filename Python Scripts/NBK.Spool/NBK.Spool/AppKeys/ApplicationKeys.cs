using System;
using System.Collections.Generic;
using System.Web;

/// <summary>
/// This class should be used for defining keys that will be used for accessing general application
/// keys. This should be considered the central place for defining these keys so developers understand
/// what is available for use.
/// </summary>
public sealed class AppKeys
{
    private AppKeys() { }

    // Part Types
    public const string SpoolFitting = "Pipe Fitting";
    public const string SpoolPipe = "Pipe";
    
    // Parameters
    public const string SpoolParameterName = "Spool Number";
    public const string SpoolItemParameterName = "Spool Item";
    public const string SpoolTypeParameterName = "Spool Part Type";
    public const string SpoolAdjacentParameterName = "Spool Adjacent";
    public const string SpoolAdjacent1ParameterName = "Spool Adjacent 1";
    public const string SpoolAdjacent2ParameterName = "Spool Adjacent 2";
    public const string SpoolConn0ParameterName = "Spool Connector 1 Name";
    public const string SpoolConn1ParameterName = "Spool Connector 2 Name";

    // Families
    public const string SpoolTitleblock = "11x17 Spool Template_NBK";
    public const string SpoolPipeTagName = "Tag Fabrication Pipe Spool Item";
    public const string SpoolFittingTagName = "Tag Fabrication Fitting Spool Item";
    public const string SpoolNorthArrow = "North Arrow";
    public const string SpoolNorthArrowIso = "North Arrow Iso";
    public const string SpoolNorthArrowModel = "North Arrow Generic Model";

    // View Types
    public const string SpoolViewPlanType = "Spool Plan";
    public const string SpoolViewSectionType = "Spool Section";
    public const string SpoolView3DType = "Spool Iso";
    public const string SpoolFittingScheduleType = "Spool Pipe Fitting Schedule";
    public const string SpoolPipeScheduleType = "Spool Pipe Schedule";
    public const string SpoolWeightScheduleType = "Spool Weight Schedule";

    // View Templates
    public const string SpoolFittingSchedViewTemplate = "Spool Pipe Fitting Schedule";
    public const string SpoolPipeSchedViewTemplate = "Spool Pipe Schedule";
    public const string SpoolWeightSchedViewTemplate = "Spool Weight Schedule";

    // Viewport Types
    public const string SpoolViewportPlanType = "Spool Plan";
    public const string SpoolViewport3DType = "Spool Iso";
    public const string SpoolViewportSectionType = "Spool Section";
    
}