package tablescanner

import (
	"./excelformat"
	"archive/zip"
	"encoding/xml"
	"fmt"
	"github.com/technix86/golang-tablescanner/excelformat"
	"io"
	"path/filepath"
	"strconv"
	"strings"
)

type xlsxStream struct {
	sheets                 []TableSheetInfo
	sheetSelected          int                  // default-opening sheet id
	iteratorPreviousRowNum int                  // last row number that Scan() have reported
	iteratorScannedRowNum  int                  // current row number fetched by reading, starting with 1
	iteratorSheetId        int                  // current row-iterating sheet id
	iteratorStream         io.ReadCloser        // current row-iterating xml stream
	iteratorData           []string             // current row-iterating row data
	iteratorDecoder        *xml.Decoder         // statefull decoder object for iterator
	iteratorXMLSegment     tIteratorXMLSegment  // current decoder xml tree location
	iteratorCapacity       int                  // default result slice capacity, synchronizes while Scan()
	zFileName              string               // original filename
	zPathSharedStrings     string               // sharedStrings.xml path from *.rels file
	zPathStyles            string               // sharedStrings.xml path from *.rels file
	z                      *zip.ReadCloser      // root zip handler
	zFiles                 map[string]*zip.File // key=zipPath
	relations              map[string]string
	referenceTable         []string
	styleNumberFormat      map[int]*excelformat.ParsedNumberFormat
	date1904               bool
	discardFormatting      bool
	discardScientific      bool
}

type tIteratorXMLSegment byte

const (
	sheetStateHidden     = "hidden"
	sheetStateVeryHidden = "veryHidden"
)

// usefull xml path checkpoints
const (
	iteratorSegmentRoot   tIteratorXMLSegment = iota // /
	iteratorSegmentW                                 // /worksheet
	iteratorSegmentWS                                // /worksheet/sheetData
	iteratorSegmentWSR                               // /worksheet/sheetData/row
	iteratorSegmentWSRC                              // /worksheet/sheetData/row/c
	iteratorSegmentWSRCIs                            // /worksheet/sheetData/row/c/is
)

func (xlsx *xlsxStream) Close() error {
	return xlsx.z.Close()
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
	defer rc.Close()
	decoder := xml.NewDecoder(rc)
	err = decoder.Decode(rels)
	if err != nil {
		return err
	}
	xlsx.zPathSharedStrings = "xl/sharedStrings.xml";
	xlsx.zPathStyles = "xl/styles.xml";
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
	defer rc.Close()
	decoder := xml.NewDecoder(rc)
	err = decoder.Decode(workbook)
	if err != nil {
		return err
	}
	xlsx.date1904 = workbook.WorkbookPr.Date1904
	xlsx.sheets = make([]TableSheetInfo, len(workbook.Sheets.Sheet))
	for idx, sheet := range workbook.Sheets.Sheet {
		// undefined path isn't critical, broken sheet can be softly ignored while fetching
		xlsx.sheets[idx] = TableSheetInfo{Name: sheet.Name, HideLevel: TableSheetVisible, path: xlsx.relations[sheet.Id], rId: sheet.Id}
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
	xlsx.SetSheetId(xlsx.sheetSelected)
	return nil
}

func (xlsx *xlsxStream) readStyles() error {
	path := xlsx.zPathStyles
	xlsx.styleNumberFormat = make(map[int]*excelformat.ParsedNumberFormat)
	parsedNumberFormatCache := make(map[int]*excelformat.ParsedNumberFormat)
	z, err := xlsx.findZipHandler(path)
	if nil != err {
		// non-critical error: styles file not found
		return nil
	}
	rc, err := z.Open()
	if err != nil {
		return err
	}
	defer rc.Close()
	decoder := xml.NewDecoder(rc)
	styles := &xmlStyleSheet{}
	err = decoder.Decode(styles)
	if err != nil {
		return err
	}
	for numFmtId, numFmt := range excelformat.BuiltInNumFmt {
		parsedNumberFormatCache[numFmtId] = excelformat.ParseNumFmt(numFmt)
	}
	for _, numFmt := range styles.NumFmts.NumFmt {
		parsedNumberFormatCache[numFmt.NumFmtId] = excelformat.ParseNumFmt(numFmt.FormatCode)
	}
	for styleId, xf := range styles.CellXfs.Xf {
		xlsx.styleNumberFormat[styleId] = parsedNumberFormatCache[xf.NumFmtId]
	}
	return nil
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
	defer rc.Close()
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
	xlsx.SetSheetId(id)
	return nil
}

func (xlsx *xlsxStream) GetSheets() []TableSheetInfo {
	result := make([]TableSheetInfo, len(xlsx.sheets))
	copy(result, xlsx.sheets)
	return result
}
func (xlsx *xlsxStream) GetCurrentSheetId() int {
	return xlsx.iteratorSheetId
}

func (xlsx *xlsxStream) SetFormatRaw() {
	xlsx.discardFormatting = true
	xlsx.discardScientific = false
}
func (xlsx *xlsxStream) SetFormatFormatted() {
	xlsx.discardFormatting = false
	xlsx.discardScientific = false
}

func (xlsx *xlsxStream) SetFormatFormattedSciFix() {
	xlsx.discardFormatting = false
	xlsx.discardScientific = true
}

func (xlsx *xlsxStream) SetSheetId(id int) error {
	xlsx.iteratorCapacity = 0
	xlsx.iteratorPreviousRowNum = 0
	xlsx.iteratorScannedRowNum = 0
	xlsx.iteratorData = []string{}
	xlsx.iteratorXMLSegment = iteratorSegmentRoot
	if nil != xlsx.iteratorStream {
		xlsx.iteratorStream.Close()
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
	if xlsx.iteratorScannedRowNum < xlsx.iteratorPreviousRowNum+1 {
		// we have detected that some rows are not present
		xlsx.iteratorPreviousRowNum++
		return []string{}
	}
	xlsx.iteratorPreviousRowNum = xlsx.iteratorScannedRowNum
	return xlsx.iteratorData
}

func (xlsx *xlsxStream) Scan() error {
	var err error
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
	currentColumnNum := 0 // 0=explicit invalid state, 1-based
	currentCellStyleStr := ""
	currentCellStyleId := -1
	currentCellTypeStr := ""
	currentCellString := ""
	rowIsParsed := false
	for !rowIsParsed {
		tok, tokenErr := xlsx.iteratorDecoder.Token()
		if tokenErr != nil || tok == nil {
			xlsx.SetSheetId(xlsx.iteratorSheetId)
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
					if !xlsx.discardFormatting {
						parsedFormat := xlsx.styleNumberFormat[currentCellStyleId]
						if nil == parsedFormat {
							// style[#currentCellStyleId].numFmt is incorrect
						} else {
							currentCellStringFormatted, err := parsedFormat.FormatValue(currentCellString, currentCellTypeStr, xlsx.date1904, xlsx.discardScientific)
							if nil != err {
								// extra virg^W error type... they haunt me
							} else {
								currentCellString = currentCellStringFormatted
							}
						}
					}
					if len(xlsx.iteratorData) >= currentColumnNum {
						xlsx.iteratorData[currentColumnNum-1] = currentCellString
					} else {
						for len(xlsx.iteratorData) < currentColumnNum-1 {
							xlsx.iteratorData = append(xlsx.iteratorData, "")
						}
						xlsx.iteratorData = append(xlsx.iteratorData, currentCellString)
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
							_, dimensions := findXmlTokenAttrValue(&tok, "ref")
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
						xlsx.iteratorData = make([]string, 0, xlsx.iteratorCapacity)
						err, currentRowNumStr := findXmlTokenAttrValue(&tok, "r")
						if nil == err {
							xlsx.iteratorScannedRowNum, err = strconv.Atoi(currentRowNumStr)
						}
						if nil != err {
							break SkipCurrentToken
						}
					}
				case iteratorSegmentWSR:
					if tok.Name.Local == "c" {
						nextSegment = iteratorSegmentWSRC
						currentCellString = ""
						currentColumnNum = -1
						_, currentCellTypeStr = findXmlTokenAttrValue(&tok, "t")
						_, currentCellStyleStr = findXmlTokenAttrValue(&tok, "s")
						currentCellStyleId, err = strconv.Atoi(currentCellStyleStr)
						if nil != err {
							currentCellStyleId = -1
						}
						err, coords := findXmlTokenAttrValue(&tok, "r")
						var currentRowNum int
						if nil == err {
							err, currentColumnNum, currentRowNum = extractCellCoords(coords)
						}
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
								xlsx.SetSheetId(xlsx.iteratorSheetId)
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

func findXmlTokenAttrValue(tok *xml.StartElement, attrName string) (error, string) {
	for _, attr := range tok.Attr {
		if attr.Name.Local == attrName {
			return nil, attr.Value
		}
	}
	return fmt.Errorf("token attr [%s] not found", attrName), ""
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
