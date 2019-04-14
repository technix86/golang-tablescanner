package tablescanner

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"path/filepath"
	"strconv"
	"strings"
)

type xlsxTableSheetInfo struct {
	Name      string
	HideLevel TSheetHideLevel
	path      string
	rId       string
}

type xlsxStream struct {
	formatter              excelFormatter
	i18n                   *tI18n   // reference to selected i18n config
	fmtI18n                []string // excel built-in number formats depending on system locale
	sheets                 []*xlsxTableSheetInfo
	sheetSelected          int                  // default-opening sheet id
	iteratorLastError      error                // error which caused last Scan() failed
	iteratorRowNum         int                  // row number that Scan() implies
	iteratorScannedRowNum  int                  // current row number fetched by reading, starting with 1
	iteratorScannedData    []string             // current row-iterating row data
	iteratorSheetId        int                  // current row-iterating sheet id
	iteratorStream         io.ReadCloser        // current row-iterating xml stream
	iteratorDecoder        *xml.Decoder         // statefull decoder object for iterator
	iteratorXMLSegment     tIteratorXMLSegment  // current decoder xml tree location
	iteratorCapacity       int                  // default result slice capacity, synchronizes while Scan()
	zFileName              string               // original filename
	zPathSharedStrings     string               // sharedStrings.xml path from *.rels file
	zPathStyles            string               // styles.xml path from *.rels file
	z                      *zip.ReadCloser      // root zip handler
	zFiles                 map[string]*zip.File // key=zipPath
	relations              map[string]string    // workbook-relation-id to path
	referenceTable         []string             // sharedStrings
	numFmtCustom           []string
	style2numFmtId         []int
	styleNumberFormatCache []*parsedNumberFormat // style-id to parsedNumberFormat
}

type tIteratorXMLSegment byte

const (
	sheetStateHidden     = "hidden"
	sheetStateVeryHidden = "veryHidden"
)

// xml path checkpoints
const (
	iteratorSegmentRoot   tIteratorXMLSegment = iota // /
	iteratorSegmentW                                 // /worksheet
	iteratorSegmentWS                                // /worksheet/sheetData
	iteratorSegmentWSR                               // /worksheet/sheetData/row
	iteratorSegmentWSRC                              // /worksheet/sheetData/row/c
	iteratorSegmentWSRCIs                            // /worksheet/sheetData/row/c/is
)

type xmlWorkbook struct {
	WorkbookPr xmlWorkbookPr `xml:"workbookPr"`
	BookViews  xmlBookViews  `xml:"bookViews"`
	Sheets     xmlSheets     `xml:"sheets"`
}

type xmlWorkbookPr struct {
	Date1904 bool `xml:"date1904,attr"`
}

type xmlBookViews struct {
	WorkBookView []xmlWorkBookView `xml:"workbookView"`
}

type xmlWorkBookView struct {
	ActiveTab int `xml:"activeTab,attr,omitempty"`
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
	Relationships []xmlWorkbookRelation `xml:"Relationship"`
}

type xmlWorkbookRelation struct {
	Id     string `xml:",attr"`
	Target string `xml:",attr"`
	Type   string `xml:",attr"`
}

type xmlStyleSheet struct {
	CellXfs xmlCellXfs `xml:"cellXfs,omitempty"`
	NumFmts xmlNumFmts `xml:"numFmts,omitempty"`
}

type xmlCellXfs struct {
	Xf []xmlXf `xml:"xf,omitempty"`
}

type xmlXf struct {
	NumFmtId int `xml:"numFmtId,attr"`
}

type xmlNumFmts struct {
	NumFmt []xmlNumFmt `xml:"numFmt,omitempty"`
}

type xmlNumFmt struct {
	NumFmtId   int    `xml:"numFmtId,attr,omitempty"`
	FormatCode string `xml:"formatCode,attr,omitempty"`
}

func newXLSXStream(fileName string) (error, ITableDocumentScanner) {
	var err error
	xlsx := &xlsxStream{zFileName: fileName}
	err = xlsx.SetI18n("en")
	if nil != err {
		return err, nil
	}
	xlsx.z, err = zip.OpenReader(fileName)
	if err != nil {
		return err, nil
	}
	xlsx.zFiles = make(map[string]*zip.File, len(xlsx.z.File))
	for _, v := range xlsx.z.File {
		xlsx.zFiles[v.Name] = v
	}
	err = xlsx.getWorkbookRelations("xl/_rels/workbook.xml.rels")
	if err != nil {
		return err, nil
	}
	err = xlsx.readSharedStrings()
	if err != nil {
		return err, nil
	}
	err = xlsx.readStyles()
	if err != nil {
		return err, nil
	}
	err = xlsx.readWorkbook("xl/workbook.xml")
	if err != nil {
		return err, nil
	}
	return nil, xlsx
}

func (sheet *xlsxTableSheetInfo) GetName() string {
	return sheet.Name
}

func (sheet *xlsxTableSheetInfo) GetHideLevel() TSheetHideLevel {
	return sheet.HideLevel
}

func (xlsx *xlsxStream) Close() error {
	return xlsx.z.Close()
}

func (sheet *xlsxStream) FormatterAvailable() bool {
	return false
}

func (xlsx *xlsxStream) Formatter() IExcelFormatter {
	return &xlsx.formatter
}

func (xlsx *xlsxStream) SetI18n(code string) error {
	if _, ok := numFmtI18n[code]; !ok {
		return fmt.Errorf("Unknown i18n[%s]", code)
	}
	xlsx.fmtI18n = []string{}
	for id, numFmt := range numFmtI18n[strings.ToLower(code)].numFmtDefaults {
		for len(xlsx.fmtI18n) < id+1 {
			xlsx.fmtI18n = append(xlsx.fmtI18n, "")
		}
		xlsx.fmtI18n[id] = numFmt
	}
	xlsx.i18n = numFmtI18n[code]
	xlsx.formatter.setI18n(xlsx.i18n)
	xlsx.styleNumberFormatCache = []*parsedNumberFormat{}
	return nil
}

func (xlsx *xlsxStream) findZipHandler(path string) (*zip.File, error) {
	z, ok := xlsx.zFiles[path]
	if ok {
		return z, nil
	}

	for pathTry, z := range xlsx.zFiles {
		if strings.ToLower(pathTry) == strings.ToLower(path) {
			// case-invalid files are possible
			return z, nil
		}
	}
	return nil, fmt.Errorf("cannot find required file %s", path)
}

func (xlsx *xlsxStream) getWorkbookRelations(path string) error {
	currentWorkbookPath := filepath.FromSlash(filepath.Dir(filepath.Dir(path)))
	defaultPathPrefix := ""
	if len(currentWorkbookPath) > 0 {
		defaultPathPrefix = currentWorkbookPath + "/"
	}
	rels := new(xmlWorkbookRels)
	xlsx.relations = make(map[string]string)
	z, err := xlsx.findZipHandler(path)
	if nil != err {
		return err
	}
	rc, err := z.Open()
	if err != nil {
		return err
	}
	defer nowarnCloseCloser(rc)
	decoder := xml.NewDecoder(rc)
	err = decoder.Decode(rels)
	if err != nil {
		return err
	}
	xlsx.zPathSharedStrings = "xl/sharedStrings.xml"
	xlsx.zPathStyles = "xl/styles.xml"
	for _, relation := range rels.Relationships {
		if relation.Target[0] == '/' {
			xlsx.relations[relation.Id] = relation.Target[1:]
		} else {
			xlsx.relations[relation.Id] = defaultPathPrefix + relation.Target
		}
		switch strings.ToLower(filepath.Base(relation.Type)) {
		case "styles":
			xlsx.zPathStyles = xlsx.relations[relation.Id]
		case "sharedstrings":
			xlsx.zPathSharedStrings = xlsx.relations[relation.Id]
		}
	}
	return nil
}
func (xlsx *xlsxStream) readWorkbook(path string) error {
	workbook := new(xmlWorkbook)
	z, err := xlsx.findZipHandler(path)
	if nil != err {
		return err
	}
	rc, err := z.Open()
	if err != nil {
		return err
	}
	defer nowarnCloseCloser(rc)
	decoder := xml.NewDecoder(rc)
	err = decoder.Decode(workbook)
	if err != nil {
		return err
	}
	xlsx.formatter = *newExcelFormatter("en")
	xlsx.formatter.setDate1904(workbook.WorkbookPr.Date1904)
	xlsx.sheets = make([]*xlsxTableSheetInfo, len(workbook.Sheets.Sheet))
	for idx, sheet := range workbook.Sheets.Sheet {
		// undefined path isn't critical, broken sheet can be softly ignored while fetching
		xlsx.sheets[idx] = &xlsxTableSheetInfo{Name: sheet.Name, HideLevel: TableSheetVisible, path: xlsx.relations[sheet.Id], rId: sheet.Id}
		if sheet.State == sheetStateHidden {
			xlsx.sheets[idx].HideLevel = TableSheetHidden
		}
		if sheet.State == sheetStateVeryHidden {
			xlsx.sheets[idx].HideLevel = TableSheetVeryHidden
		}
	}
	if len(workbook.BookViews.WorkBookView) > 0 {
		xlsx.sheetSelected = workbook.BookViews.WorkBookView[0].ActiveTab
		if xlsx.sheetSelected > len(xlsx.sheets)-1 {
			xlsx.sheetSelected = len(xlsx.sheets) - 1
		}
		if xlsx.sheetSelected < 0 {
			xlsx.sheetSelected = 0
		}
	}
	_ = xlsx.SetSheetId(xlsx.sheetSelected)
	return nil
}

func (xlsx *xlsxStream) readStyles() error {
	path := xlsx.zPathStyles
	/*
		numFmtCustom           []string
		style2numFmtId         []int
		styleNumberFormatCache []*parsedNumberFormat // style-id to parsedNumberFormat
	*/
	xlsx.numFmtCustom = make([]string, 0, 256)
	xlsx.style2numFmtId = make([]int, 0, 32)
	xlsx.styleNumberFormatCache = make([]*parsedNumberFormat, 0, 256)
	z, err := xlsx.findZipHandler(path)
	if nil != err {
		// non-critical error: styles file not found
		return nil
	}
	rc, err := z.Open()
	if err != nil {
		return err
	}
	defer nowarnCloseCloser(rc)
	decoder := xml.NewDecoder(rc)
	styles := &xmlStyleSheet{}
	err = decoder.Decode(styles)
	if err != nil {
		return err
	}
	for _, numFmt := range styles.NumFmts.NumFmt {
		for len(xlsx.numFmtCustom) < numFmt.NumFmtId+1 {
			xlsx.numFmtCustom = append(xlsx.numFmtCustom, "")
		}
		/*
			if len(numFmt.FormatCode) >= 2 && numFmt.FormatCode[0] == '[' && numFmt.FormatCode[1] == '$' {
				SystemRefEnd := strings.IndexRune(numFmt.FormatCode, ']')
				if SystemRefEnd >= 0 {
					numFmt.FormatCode = numFmt.FormatCode[0:SystemRefEnd+1]
				}
			}
		*/
		xlsx.numFmtCustom[numFmt.NumFmtId] = numFmt.FormatCode
	}
	for styleId, xf := range styles.CellXfs.Xf {
		for len(xlsx.style2numFmtId) < styleId+1 {
			xlsx.style2numFmtId = append(xlsx.style2numFmtId, 0)
		}
		xlsx.style2numFmtId[styleId] = xf.NumFmtId
	}
	return nil
}

// number formats are parsed only when needed
// it guarantees that parser parameter i18n affects caches only while scanning table
func (xlsx *xlsxStream) getParsedNumFmtByStyle(styleId int) *parsedNumberFormat {
	if styleId >= 0 && styleId < len(xlsx.styleNumberFormatCache) { // if inside cached interval
		if nil != xlsx.styleNumberFormatCache[styleId] { // search in cache
			return xlsx.styleNumberFormatCache[styleId]
		}
	}
	if len(xlsx.style2numFmtId) == 0 {
		// maybe panic?
		// we have to choose style from empty set
		xlsx.style2numFmtId = []int{0} // make default style with "general" fmt
	}
	if len(xlsx.numFmtCustom) == 0 { // numFmtCustom cannot be empty and must have at least len(builtin) items
		xlsx.numFmtCustom = make([]string, len(xlsx.fmtI18n))
	}
	if styleId < 0 || styleId >= len(xlsx.style2numFmtId) {
		// maybe panic again?
		// if outside valid id interval use first known style
		return xlsx.getParsedNumFmtByStyle(0)
	}
	numFmtId := xlsx.style2numFmtId[styleId]
	if numFmtId < 0 || numFmtId >= len(xlsx.numFmtCustom) {
		numFmtId = 0
	}
	var numFmt string
	if numFmtId < len(xlsx.fmtI18n) && "" == xlsx.numFmtCustom[numFmtId] {
		numFmt = xlsx.fmtI18n[numFmtId]
	} else {
		numFmt = xlsx.numFmtCustom[numFmtId]
	}
	for len(xlsx.styleNumberFormatCache) < styleId+1 {
		xlsx.styleNumberFormatCache = append(xlsx.styleNumberFormatCache, nil)
	}
	if len(numFmt) >= 2 && numFmt[0] == '[' && numFmt[1] == '$' {
		SystemRefEnd := strings.IndexRune(numFmt, ']')
		if SystemRefEnd >= 0 {
			if systemFmt, found := xlsx.i18n.numFmtSystem[numFmt[0:SystemRefEnd+1]]; found {
				numFmt = systemFmt
			} else {
				numFmt = numFmt[SystemRefEnd+1:]
			}
		}
	}
	xlsx.styleNumberFormatCache[styleId] = parseNumFmt(numFmt)
	return xlsx.styleNumberFormatCache[styleId]
}

func (xlsx *xlsxStream) readSharedStrings() error {
	path := xlsx.zPathSharedStrings
	z, err := xlsx.findZipHandler(path)
	if nil != err {
		// non-critical error: sharedStrings file not found
		return nil
	}
	rc, err := z.Open()
	if err != nil {
		return err
	}
	defer nowarnCloseCloser(rc)
	decoder := xml.NewDecoder(rc)
	var stateStr string
	var tmp string
	for {
		tok, tokenErr := decoder.Token()
		if tokenErr == io.EOF {
			break
		} else if tokenErr != nil {
			return tokenErr
		}
		if tok == nil {
			fmt.Println("t is nil break")
		}
		switch tok := tok.(type) {
		case xml.EndElement:
			if tok.Name.Local == "si" {
				xlsx.referenceTable = append(xlsx.referenceTable, stateStr)
			}
		case xml.StartElement:
			if tok.Name.Local == "si" {
				stateStr = ""
			}
			if tok.Name.Local == "t" {
				if err := decoder.DecodeElement(&tmp, &tok); err != nil {
					return err
				}
				stateStr = stateStr + tmp
			}
		}
	}
	return nil
}

func (xlsx *xlsxStream) SwitchSheet(id int) error {
	if id < 0 || id >= len(xlsx.sheets) {
		return fmt.Errorf("sheet id is out of range")
	}
	_ = xlsx.SetSheetId(id)
	return nil
}

func (xlsx *xlsxStream) GetSheets() []ITableSheetInfo {
	res := make([]ITableSheetInfo, len(xlsx.sheets))
	for i, sheet := range xlsx.sheets {
		res[i] = sheet
	}
	return res
}

func (xlsx *xlsxStream) GetCurrentSheetId() int {
	return xlsx.iteratorSheetId
}

func (xlsx *xlsxStream) SetSheetId(id int) error {
	xlsx.iteratorLastError = nil
	xlsx.iteratorCapacity = 0
	xlsx.iteratorRowNum = 0
	xlsx.iteratorScannedRowNum = 0
	xlsx.iteratorScannedData = []string{}
	xlsx.iteratorXMLSegment = iteratorSegmentRoot
	if nil != xlsx.iteratorStream {
		_ = xlsx.iteratorStream.Close()
		xlsx.iteratorStream = nil // force rewind
	}
	if id < 0 || id > len(xlsx.sheets) {
		return fmt.Errorf("sheet #%d not found", id)
	}
	_, err := xlsx.findZipHandler(xlsx.sheets[id].path)
	if nil != err {
		return err
	}
	xlsx.iteratorSheetId = id
	return nil
}
func (xlsx *xlsxStream) GetScanned() []string {
	if xlsx.iteratorScannedRowNum > xlsx.iteratorRowNum {
		return []string{}
	}
	return xlsx.iteratorScannedData
}

func (xlsx *xlsxStream) GetLastScanError() error {
	return xlsx.iteratorLastError
}

func (xlsx *xlsxStream) requireScanStream() error {
	if nil == xlsx.iteratorStream {
		z, err := xlsx.findZipHandler(xlsx.sheets[xlsx.iteratorSheetId].path)
		if nil != err {
			return fmt.Errorf("sheet #%d not found: %s", xlsx.iteratorSheetId, err)
		}
		xlsx.iteratorStream, err = z.Open()
		if err != nil {
			return fmt.Errorf("file stream [%s] Open() failed: %s", xlsx.sheets[xlsx.iteratorSheetId].path, err.Error())
		}
		xlsx.iteratorDecoder = xml.NewDecoder(xlsx.iteratorStream)
	}
	return nil
}

func (xlsx *xlsxStream) Scan() (err error) {
	// if row we have scanned is not next to previously returned, just increase "previouslyReturned" counter and imply empty row
	if xlsx.iteratorScannedRowNum > xlsx.iteratorRowNum {
		xlsx.iteratorRowNum++
	} else {
		err = xlsx.scanInternal()
		if nil == err {
			xlsx.iteratorRowNum++
		}
	}
	xlsx.iteratorLastError = err
	return err
}

func (xlsx *xlsxStream) scanInternal() (err error) {
	err = xlsx.requireScanStream()
	if nil != err {
		return err
	}
	currentColumnNum := 0 // 0=explicit invalid state, 1-based
	currentCellStyleStr := ""
	currentCellStyleId := -1
	currentCellTypeStr := ""
	currentCellString := ""
	rowIsParsed := false
	for !rowIsParsed {
		tok, tokenErr := xlsx.iteratorDecoder.Token()
		if tokenErr != nil || tok == nil {
			_ = xlsx.SetSheetId(xlsx.iteratorSheetId)
			if io.EOF == tokenErr {
				return tokenErr
			}
			return fmt.Errorf("xml token read error in [%s] at pos %d: %s", xlsx.sheets[xlsx.iteratorSheetId].path, xlsx.iteratorDecoder.InputOffset(), tokenErr.Error())
		}
		switch tok := tok.(type) {
		case xml.EndElement:
			switch xlsx.iteratorXMLSegment {
			case iteratorSegmentW:
				if tok.Name.Local == "worksheet" {
					xlsx.iteratorXMLSegment = iteratorSegmentRoot
				}
			case iteratorSegmentWS:
				if tok.Name.Local == "sheetData" {
					xlsx.iteratorXMLSegment = iteratorSegmentW
				}
			case iteratorSegmentWSR:
				if tok.Name.Local == "row" {
					xlsx.iteratorXMLSegment = iteratorSegmentWS
					rowIsParsed = true
				}
			case iteratorSegmentWSRC:
				if tok.Name.Local == "c" {
					xlsx.iteratorXMLSegment = iteratorSegmentWSR
					if currentColumnNum < 1 {
						panic(fmt.Sprintf("WTF i'm doing here? Cell have to been skipped in this condition! [file=%s sheet=%s at pos %d]", xlsx.zFileName, xlsx.sheets[xlsx.iteratorSheetId].path, xlsx.iteratorDecoder.InputOffset()))
					}
					parsedFormat := xlsx.getParsedNumFmtByStyle(currentCellStyleId)
					if nil == parsedFormat {
						// style[#currentCellStyleId].numFmt is incorrect
					} else {
						currentCellStringFormatted, err := xlsx.formatter.FormatValue(currentCellString, currentCellTypeStr, parsedFormat)
						if nil != err {
							// extra virg^W error type... they haunt me
						} else {
							currentCellString = currentCellStringFormatted
						}
					}
					if len(xlsx.iteratorScannedData) >= currentColumnNum {
						xlsx.iteratorScannedData[currentColumnNum-1] = currentCellString
					} else {
						for len(xlsx.iteratorScannedData) < currentColumnNum-1 {
							xlsx.iteratorScannedData = append(xlsx.iteratorScannedData, "")
						}
						xlsx.iteratorScannedData = append(xlsx.iteratorScannedData, currentCellString)
					}
				}
			case iteratorSegmentWSRCIs:
				if tok.Name.Local == "is" {
					xlsx.iteratorXMLSegment = iteratorSegmentWSRC
				}
			}
		case xml.StartElement:
			var skipToken = true
			tagIsDecoded := false
			nextSegment := xlsx.iteratorXMLSegment
		SkipCurrentToken:
			for { // single-cycle loop is used to break the logic anywhere (I do not like closures here) and fully skip current tag
				switch xlsx.iteratorXMLSegment {
				case iteratorSegmentRoot:
					if tok.Name.Local == "worksheet" {
						nextSegment = iteratorSegmentW
					} else {
						if tok.Name.Local == "dimension" {
							dimensions,_ := findXmlTokenAttrValue(&tok, "ref")
							_, _, _, xCoordMax, _ := extractCellRangeCoords(dimensions)
							xlsx.iteratorCapacity = xCoordMax
						}
						break SkipCurrentToken
					}

				case iteratorSegmentW:
					if tok.Name.Local == "sheetData" {
						nextSegment = iteratorSegmentWS
					}
				case iteratorSegmentWS:
					if tok.Name.Local == "row" {
						nextSegment = iteratorSegmentWSR
						xlsx.iteratorScannedData = make([]string, 0, xlsx.iteratorCapacity)
						currentRowNumStr,attrExists := findXmlTokenAttrValue(&tok, "r")
						if attrExists {
							// attr "r" present, require valid int and greater than previous value
							xlsx.iteratorScannedRowNum, err = strconv.Atoi(currentRowNumStr)
							if nil != err {
								return fmt.Errorf("row number \"%s\" is not int file=%s pos=%d", currentRowNumStr, xlsx.sheets[xlsx.iteratorSheetId].path, xlsx.iteratorDecoder.InputOffset())
							}
							if xlsx.iteratorScannedRowNum <= xlsx.iteratorRowNum {
								return fmt.Errorf("row numbers are not strictly increasing ...%d...%d... file=%s pos=%d", xlsx.iteratorRowNum, xlsx.iteratorScannedRowNum, xlsx.sheets[xlsx.iteratorSheetId].path, xlsx.iteratorDecoder.InputOffset())
							}
						} else {
							xlsx.iteratorScannedRowNum = xlsx.iteratorRowNum + 1
						}
					}
				case iteratorSegmentWSR:
					if tok.Name.Local == "c" {
						nextSegment = iteratorSegmentWSRC
						currentCellString = ""
						currentColumnNum = -1
						currentCellTypeStr,_ = findXmlTokenAttrValue(&tok, "t")
						currentCellStyleStr,_ = findXmlTokenAttrValue(&tok, "s")
						currentCellStyleId, err = strconv.Atoi(currentCellStyleStr)
						if nil != err {
							currentCellStyleId = -1
						}
						coords,ok := findXmlTokenAttrValue(&tok, "r")
						var currentRowNum int
						if !ok {
							break SkipCurrentToken
						}
						err, currentColumnNum, currentRowNum = extractCellCoords(coords)
						if nil == err && currentRowNum != xlsx.iteratorScannedRowNum {
							err = fmt.Errorf("row[%d] != cell.row[%d] for cell %s", currentRowNum, xlsx.iteratorScannedRowNum, coords)
						}
						if nil != err {
							break SkipCurrentToken
						}
						if currentColumnNum > xlsx.iteratorCapacity {
							xlsx.iteratorCapacity = currentColumnNum
						}
					}
				case iteratorSegmentWSRC:
					if tok.Name.Local == "is" {
						if currentCellTypeStr != "inlineStr" {
							// error: <is> tags requires <c t=inlineStr>
							break SkipCurrentToken
						}
						nextSegment = iteratorSegmentWSRCIs
					} else if currentCellTypeStr != "inlineStr" {
						if tok.Name.Local == "v" {
							var tagValue string
							err = xlsx.iteratorDecoder.DecodeElement(&tagValue, &tok)
							tagIsDecoded = true
							if nil != err {
								// string decoding failed
								_ = xlsx.SetSheetId(xlsx.iteratorSheetId)
								return fmt.Errorf("xml string decoding error in [%s] at pos %d: %s", xlsx.sheets[xlsx.iteratorSheetId].path, xlsx.iteratorDecoder.InputOffset(), err.Error())
							}
							if currentCellTypeStr == "s" { // type = shared strings
								strId, err := strconv.Atoi(strings.Trim(tagValue, " "))
								if nil == err {
									if strId < 0 || strId >= len(xlsx.referenceTable) {
										// invalid string index
										tagValue = ""
										break SkipCurrentToken
									}
									tagValue = xlsx.referenceTable[strId]
								}
							}
							currentCellString += tagValue
						} else {
							break SkipCurrentToken
						}
					}
				case iteratorSegmentWSRCIs:
					switch tok.Name.Local {
					case "r":
						// do nothing, fetch next tag
					case "t":
						if currentColumnNum < 1 {
							panic(fmt.Sprintf("WTF i'm doing here? Cell have to been skipped in this condition! [file=%s sheet=%s at pos %d]", xlsx.zFileName, xlsx.sheets[xlsx.iteratorSheetId].path, xlsx.iteratorDecoder.InputOffset()))
						}
						var tagValue string
						err = xlsx.iteratorDecoder.DecodeElement(&tagValue, &tok)
						tagIsDecoded = true
						if nil != err {
							// string decoding failed
							break SkipCurrentToken
						}
						currentCellString += tagValue
					default:
						break SkipCurrentToken
					}
				}
				skipToken = false
				break

			}
			if tagIsDecoded {

			} else if skipToken {
				_ = xlsx.iteratorDecoder.Skip()
			} else {
				xlsx.iteratorXMLSegment = nextSegment
			}
		}
	}
	return nil
}

// @todo: implement simple ok or not ok instead of error
func findXmlTokenAttrValue(tok *xml.StartElement, attrName string) (string,bool) {
	for _, attr := range tok.Attr {
		if attr.Name.Local == attrName {
			return attr.Value,true
		}
	}
	return "",false
}

func extractCellRangeCoords(cellRangeAddr string) (err error, x1 int, y1 int, x2 int, y2 int) {
	l := strings.Split(cellRangeAddr, ":")
	if len(l) != 2 {
		return fmt.Errorf("invalid cell range syntax (%s)", cellRangeAddr), 0, 0, 0, 0
	}
	err, x1, y1 = extractCellCoords(l[0])
	if nil != err {
		return
	}
	err, x2, y2 = extractCellCoords(l[1])
	return
}

// parse A5/D4/ZZ2354 coords (1-based)
func extractCellCoords(cellAddr string) (err error, x int, y int) {
	for idx, char := range cellAddr {
		switch {
		case '0' <= char && char <= '9':
			y = y*10 + int(char-'0')
		case 'A' <= char && char <= 'Z':
			x = x*26 + int(char-'A') + 1
		case 'a' <= char && char <= 'z':
			x = x*26 + int(char-'a') + 1
		default:
			return fmt.Errorf("undefined character %c at pos %d", char, idx), 0, 0
		}
	}
	return nil, x, y
}
