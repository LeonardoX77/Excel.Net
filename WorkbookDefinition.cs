using System;
using System.Xml.Serialization;
using System.Collections.Generic;


namespace Excel
{
    [Serializable()]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(ElementName = "workbook", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class WorkbookDefinition
    {
        [XmlElement(ElementName = "fileVersion", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public FileVersion FileVersion { get; set; }
        [XmlElement(ElementName = "workbookPr", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public WorkbookPr WorkbookPr { get; set; }
        [XmlElement(ElementName = "AlternateContent", Namespace = "http://schemas.openxmlformats.org/markup-compatibility/2006")]
        public AlternateContent AlternateContent { get; set; }
        [XmlElement(ElementName = "bookViews", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public BookViews BookViews { get; set; }
        [XmlElement(ElementName = "sheets", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public Sheets Sheets { get; set; }
        [XmlElement(ElementName = "definedNames", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public DefinedNames DefinedNames { get; set; }
        [XmlElement(ElementName = "calcPr", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public CalcPr CalcPr { get; set; }
        [XmlElement(ElementName = "extLst", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public ExtLst ExtLst { get; set; }
        [XmlAttribute(AttributeName = "xmlns")]
        public string Xmlns { get; set; }
        [XmlAttribute(AttributeName = "r", Namespace = "http://www.w3.org/2000/xmlns/")]
        public string R { get; set; }
        [XmlAttribute(AttributeName = "mc", Namespace = "http://www.w3.org/2000/xmlns/")]
        public string Mc { get; set; }
        [XmlAttribute(AttributeName = "Ignorable", Namespace = "http://schemas.openxmlformats.org/markup-compatibility/2006")]
        public string Ignorable { get; set; }
        [XmlAttribute(AttributeName = "x15", Namespace = "http://www.w3.org/2000/xmlns/")]
        public string X15 { get; set; }
    }


    [XmlRoot(ElementName = "fileVersion", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class FileVersion
    {
        [XmlAttribute(AttributeName = "appName")]
        public string AppName { get; set; }
        [XmlAttribute(AttributeName = "lastEdited")]
        public string LastEdited { get; set; }
        [XmlAttribute(AttributeName = "lowestEdited")]
        public string LowestEdited { get; set; }
        [XmlAttribute(AttributeName = "rupBuild")]
        public string RupBuild { get; set; }
    }

    [XmlRoot(ElementName = "workbookPr", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class WorkbookPr
    {
        [XmlAttribute(AttributeName = "defaultThemeVersion")]
        public string DefaultThemeVersion { get; set; }
    }

    [XmlRoot(ElementName = "absPath", Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac")]
    public class AbsPath
    {
        [XmlAttribute(AttributeName = "url")]
        public string Url { get; set; }
        [XmlAttribute(AttributeName = "x15ac", Namespace = "http://www.w3.org/2000/xmlns/")]
        public string X15ac { get; set; }
    }

    [XmlRoot(ElementName = "Choice", Namespace = "http://schemas.openxmlformats.org/markup-compatibility/2006")]
    public class Choice
    {
        [XmlElement(ElementName = "absPath", Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac")]
        public AbsPath AbsPath { get; set; }
        [XmlAttribute(AttributeName = "Requires")]
        public string Requires { get; set; }
    }

    [XmlRoot(ElementName = "AlternateContent", Namespace = "http://schemas.openxmlformats.org/markup-compatibility/2006")]
    public class AlternateContent
    {
        [XmlElement(ElementName = "Choice", Namespace = "http://schemas.openxmlformats.org/markup-compatibility/2006")]
        public Choice Choice { get; set; }
        [XmlAttribute(AttributeName = "mc", Namespace = "http://www.w3.org/2000/xmlns/")]
        public string Mc { get; set; }
    }

    [XmlRoot(ElementName = "workbookView", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class WorkbookView
    {
        [XmlAttribute(AttributeName = "xWindow")]
        public string XWindow { get; set; }
        [XmlAttribute(AttributeName = "yWindow")]
        public string YWindow { get; set; }
        [XmlAttribute(AttributeName = "windowWidth")]
        public string WindowWidth { get; set; }
        [XmlAttribute(AttributeName = "windowHeight")]
        public string WindowHeight { get; set; }
        [XmlAttribute(AttributeName = "activeTab")]
        public string ActiveTab { get; set; }
    }

    [XmlRoot(ElementName = "bookViews", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class BookViews
    {
        [XmlElement(ElementName = "workbookView", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public WorkbookView WorkbookView { get; set; }
    }

    [XmlRoot(ElementName = "sheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class Sheet
    {
        [XmlAttribute(AttributeName = "name")]
        public string Name { get; set; }
        [XmlAttribute(AttributeName = "sheetId")]
        public string SheetId { get; set; }
        [XmlAttribute(AttributeName = "id", Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string Id { get; set; }
    }

    [XmlRoot(ElementName = "sheets", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class Sheets
    {
        [XmlElement(ElementName = "sheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public List<Sheet> Sheet { get; set; }
    }

    [XmlRoot(ElementName = "definedName", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class DefinedName
    {
        [XmlAttribute(AttributeName = "name")]
        public string Name { get; set; }
        [XmlAttribute(AttributeName = "hidden")]
        public string Hidden { get; set; }

        private string _Text = "";
        [XmlText]
        public string Text {
            get
            {
                return _Text;
            }
            set
            {
                _Text = value.Replace("$", "");
            }
        }
        
        [XmlAttribute(AttributeName = "localSheetId")]
        public string LocalSheetId { get; set; }
    }

    [XmlRoot(ElementName = "definedNames", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class DefinedNames
    {
        [XmlElement(ElementName = "definedName", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public List<DefinedName> DefinedName { get; set; }
    }

    [XmlRoot(ElementName = "calcPr", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CalcPr
    {
        [XmlAttribute(AttributeName = "calcId")]
        public string CalcId { get; set; }
    }

    [XmlRoot(ElementName = "workbookPr", Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")]
    public class WorkbookPr2
    {
        [XmlAttribute(AttributeName = "chartTrackingRefBase")]
        public string ChartTrackingRefBase { get; set; }
    }

    [XmlRoot(ElementName = "ext", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class Ext
    {
        [XmlElement(ElementName = "workbookPr", Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")]
        public WorkbookPr2 WorkbookPr2 { get; set; }
        [XmlAttribute(AttributeName = "uri")]
        public string Uri { get; set; }
        [XmlAttribute(AttributeName = "x15", Namespace = "http://www.w3.org/2000/xmlns/")]
        public string X15 { get; set; }
    }

    [XmlRoot(ElementName = "extLst", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class ExtLst
    {
        [XmlElement(ElementName = "ext", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public Ext Ext { get; set; }
    }
}
