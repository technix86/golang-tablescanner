package tablescanner

type xmlWorkbook struct {
	//XMLName xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main workbook"`
	//FileVersion        xmlFileVersion        `xml:"fileVersion"`
	WorkbookPr xmlWorkbookPr `xml:"workbookPr"`
	//WorkbookProtection xmlWorkbookProtection `xml:"workbookProtection"`
	BookViews xmlBookViews `xml:"bookViews"`
	Sheets    xmlSheets    `xml:"sheets"`
	//DefinedNames       xmlDefinedNames       `xml:"definedNames"`
	//CalcPr             xmlCalcPr             `xml:"calcPr"`
}

type xmlBookViews struct {
	WorkBookView []xmlWorkBookView `xml:"workbookView"`
}

type xmlWorkBookView struct {
	ActiveTab int `xml:"activeTab,attr,omitempty"`
	//FirstSheet           int    `xml:"firstSheet,attr,omitempty"`
	//ShowHorizontalScroll bool   `xml:"showHorizontalScroll,attr,omitempty"`
	//ShowVerticalScroll   bool   `xml:"showVerticalScroll,attr,omitempty"`
	//ShowSheetTabs        bool   `xml:"showSheetTabs,attr,omitempty"`
	//TabRatio             int    `xml:"tabRatio,attr,omitempty"`
	//WindowHeight         int    `xml:"windowHeight,attr,omitempty"`
	//WindowWidth          int    `xml:"windowWidth,attr,omitempty"`
	//XWindow              string `xml:"xWindow,attr,omitempty"`
	//YWindow              string `xml:"yWindow,attr,omitempty"`
}

type xmlWorkbookPr struct {
	//DefaultThemeVersion string `xml:"defaultThemeVersion,attr,omitempty"`
	//BackupFile          bool   `xml:"backupFile,attr,omitempty"`
	//ShowObjects         string `xml:"showObjects,attr,omitempty"`
	Date1904 bool `xml:"date1904,attr"`
}

type xmlSheets struct {
	Sheet []xmlSheet `xml:"sheet"`
}

type xmlSheet struct {
	Name    string `xml:"name,attr,omitempty"`
	SheetId string `xml:"sheetId,attr,omitempty"`
	Id      string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
	State   string `xml:"state,attr,omitempty"`
}

type xmlWorkbookRels struct {
	//XMLName       xml.Name               `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []xmlWorkbookRelation `xml:"Relationship"`
}

type xmlWorkbookRelation struct {
	Id     string `xml:",attr"`
	Target string `xml:",attr"`
	Type   string `xml:",attr"`
}

type xmlStyleSheet struct {
	//XMLName xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main styleSheet"`
	//CellStyles   *xmlCellStyles   `xml:"cellStyles,omitempty"`
	//CellStyleXfs *xmlCellStyleXfs `xml:"cellStyleXfs,omitempty"`
	CellXfs xmlCellXfs `xml:"cellXfs,omitempty"`
	NumFmts xmlNumFmts `xml:"numFmts,omitempty"`
	//numFmtRefTable map[int]xmlNumFmt
	//parsedNumFmtTable map[string]*parsedNumberFormat
}

type xmlCellXfs struct {
	//Count int      `xml:"count,attr"`
	Xf []xmlXf `xml:"xf,omitempty"`
}

type xmlXf struct {
	//ApplyAlignment    bool          `xml:"applyAlignment,attr"`
	//ApplyBorder       bool          `xml:"applyBorder,attr"`
	//ApplyFont         bool          `xml:"applyFont,attr"`
	//ApplyFill         bool          `xml:"applyFill,attr"`
	//ApplyNumberFormat bool          `xml:"applyNumberFormat,attr"`
	//ApplyProtection   bool          `xml:"applyProtection,attr"`
	//BorderId          int           `xml:"borderId,attr"`
	//FillId            int           `xml:"fillId,attr"`
	//FontId            int           `xml:"fontId,attr"`
	NumFmtId int `xml:"numFmtId,attr"`
	//XfId     *int `xml:"xfId,attr,omitempty"`
	//Alignment         xmlAlignment `xml:"alignment"`
}

type xmlNumFmts struct {
	Count  int         `xml:"count,attr"`
	NumFmt []xmlNumFmt `xml:"numFmt,omitempty"`
}

type xmlNumFmt struct {
	NumFmtId   int    `xml:"numFmtId,attr,omitempty"`
	FormatCode string `xml:"formatCode,attr,omitempty"`
}
