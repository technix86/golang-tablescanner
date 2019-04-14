package tablescanner

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"regexp"
	"strconv"
	"strings"
	"unicode/utf16"
)

type xmlTableSheetInfo struct {
	Name      string
	HideLevel TSheetHideLevel
	start     int64 // offset of <Worksheet>
	stop      int64 // offset of </Worksheet>
}

type ReadSeekCloser interface {
	io.ReadCloser
	io.Seeker
}

type tIteratorRAWXMLSegment byte

const (
	iteratorRXSegmentRoot tIteratorRAWXMLSegment = iota // <Workbook>./
	iteratorRXSegmentW                                  // <Workbook>./Worksheet
	iteratorRXSegmentWT                                 // <Workbook>./Worksheet/Table
	iteratorRXSegmentWTR                                // <Workbook>./Worksheet/Table/Row
	iteratorRXSegmentWTRC                               // <Workbook>./Worksheet/Table/Row/Cell
)

type xmlHandle struct {
	formatter                    excelFormatter
	sheets                       []*xmlTableSheetInfo
	sheetSelected                int                    // default-opening sheet id
	iteratorLastError            error                  // error which caused last Scan() failed
	iteratorStreamSource         ReadSeekCloser         // current row-iterating xml stream
	iteratorStreamXML            io.ReadSeeker          // current row-iterating xml stream
	iteratorCapacity             int                    // default result slice capacity, synchronizes while Scan()
	iteratorDecoderInitialOffset int64                  // (real offset = init offset + decoder relative offset)
	iteratorDecoder              *xml.Decoder           // statefull decoder object for iterator
	iteratorXMLSegment           tIteratorRAWXMLSegment // current decoder xml tree location
	iteratorScannedRowNum        int                    // current row number fetched by reading, starting with 1
	iteratorScannedData          []string               // current row-iterating row data
	iteratorRowNum               int                    // row number that Scan() implies (starting with 1)
	iteratorSheetId              int                    // current row-iterating sheet id
}

type rawxmlWorksheetOptions struct {
	Visible  string     `xml:"Visible,omitempty"`  // "SheetHidden"/"SheetVeryHidden"/""
	Selected []struct{} `xml:"Selected,omitempty"` // <selected /> = []bool{false}
}

type rawxmlCell struct {
	//Type  string     `xml:"Type,attr"`  // "String"/"Number" expected only, not yet used
	Data string `xml:"Data,omitempty"`
}

func makeUTF8BufferFromUTF16(reader io.Reader, isLittleEndian bool, contentLength int64) []byte {
	resultBuf := make([]byte, 0, contentLength/2)
	chunkSize := 16384
	bufSrc := make([]byte, chunkSize)
	posByteLow := 1
	posByteHigh := 0
	chunkUTF16 := make([]uint16, chunkSize/2)
	if isLittleEndian {
		posByteLow, posByteHigh = posByteHigh, posByteLow
	}
	for chunk := 0; ; chunk++ {
		readedBytes, err := reader.Read(bufSrc)
		chunkUTF16 = chunkUTF16[0:0]
		if nil != err {
			break
		}
		for i := 0; i < readedBytes/2; i ++ {
			chunkUTF16 = append(chunkUTF16, uint16(bufSrc[i*2+posByteHigh])<<8|uint16(bufSrc[i*2+posByteLow]))
		}
		chunkUTF8 := string(utf16.Decode(chunkUTF16))
		if 0 == chunk {
			reEncReplace, _ := regexp.Compile("(?i)^(?:\xef\xbb\xbf)?(<\\?xml [^>]+)\\sencoding=\"utf-16\"")
			chunkUTF8 = reEncReplace.ReplaceAllString(chunkUTF8, "$1 encoding=\"utf-8\"")
		}
		resultBuf = append(resultBuf, []byte(chunkUTF8)...)
	}
	return resultBuf
}

func newXMLStream(fileName string, textEncoding TTextEnconding, BOMPresent []byte) (error, ITableDocumentScanner) {
	var err error
	xls := &xmlHandle{}
	fileHandle, err := os.Open(fileName)
	xls.iteratorStreamSource = fileHandle
	if err != nil {
		return err, nil
	}
	fileStat, err := fileHandle.Stat()
	fileSize := fileStat.Size()
	/* not yet supported
	err = xls.SetI18n("en")
	if err != nil {
		return err, nil
	}
	*/
	//var offsetBom int64 =0
	var xmlDecodableBuf io.ReadSeeker
	xmlDecodableBuf = xls.iteratorStreamSource
	if true {
		//_, err = xls.iteratorStreamSource.Seek(int64(len(BOMPresent)), io.SeekStart)
		_, err = xls.iteratorStreamSource.Seek(0, io.SeekStart)
		if err != nil {
			return err, nil
		}
		switch textEncoding {
		case EncodingUTF16BE:
			fallthrough
		case EncodingUTF16LE:
			bufferLength := fileSize - int64(len(BOMPresent))
			bufUtf8 := makeUTF8BufferFromUTF16(xls.iteratorStreamSource, EncodingUTF16LE == textEncoding, bufferLength)
			xmlDecodableBuf = bytes.NewReader(bufUtf8)
		case EncodingUTF8, EncodingUnknown:
			xmlDecodableBuf = xls.iteratorStreamSource
		default:
			return fmt.Errorf("text encoding of file(%s) has unservale value %#d", fileName, textEncoding), nil
		}
	}

	xls.iteratorStreamXML = xmlDecodableBuf
	_, _ = xls.iteratorStreamXML.Seek(0, io.SeekStart)
	xls.iteratorDecoder = xml.NewDecoder(xls.iteratorStreamXML)
	level := 0 // 0=/ 1=/Workbook 2=/Workbook/Worksheet
	currentSheetId := -1
	var currentSheetOptions = &rawxmlWorksheetOptions{}
	var currentSheetOpenOffset int64
	currentSheetOpenOffset = -1
	var currentSheetTableName string
	for {
		offset := xls.iteratorDecoder.InputOffset()
		tok, tokenErr := xls.iteratorDecoder.Token()
		if io.EOF == tokenErr {
			break
		}
		if tokenErr != nil || tok == nil {
			return fmt.Errorf("xml token read error at pos %d: %s", offset, tokenErr.Error()), nil
		}
		switch tok := tok.(type) {
		case xml.EndElement:
			switch tok.Name.Local {
			case "Workbook":
				if 1 == level {
					level = 0
				}
			case "Worksheet":
				if 2 == level {
					if nil == currentSheetOptions {
						return fmt.Errorf("sheet #%d has no <WorksheetOptions> section"), nil
					}
					if -1 == currentSheetOpenOffset {
						return fmt.Errorf("sheet #%d has no <Table> section"), nil
					}
					if nil != err {
						return err, nil
					}
					sheet := &xmlTableSheetInfo{
						Name:      currentSheetTableName,
						HideLevel: TableSheetVisible,
						start:     currentSheetOpenOffset,
						stop:      offset,
					}
					switch strings.ToLower(currentSheetOptions.Visible) {
					case "sheethidden":
						sheet.HideLevel = TableSheetHidden
					case "sheetveryhidden":
						sheet.HideLevel = TableSheetVeryHidden
					}
					if -1 == currentSheetId {
						return fmt.Errorf("var currentSheetId has not been initialized yet... parser FSM seems to be inconsistent"), nil
					}
					xls.sheets = append(xls.sheets, sheet)
					if 0 < len(currentSheetOptions.Selected) {
						xls.sheetSelected = currentSheetId
					}
					level = 1
				}
			}
		case xml.StartElement:
			switch tok.Name.Local {
			case "Workbook":
				if 0 == level {
					level = 1
				}
			case "Worksheet":
				if 1 == level {
					level = 2
					currentSheetId = len(xls.sheets)
					currentSheetOptions = &rawxmlWorksheetOptions{}
					currentSheetOpenOffset = offset
					_, currentSheetTableName = findXmlTokenAttrValue(&tok, "Name")
				}
			case "WorksheetOptions":
				if 2 == level {
					err = xls.iteratorDecoder.DecodeElement(currentSheetOptions, &tok)
					if nil != err {
						return fmt.Errorf("Cannot decode <WorksheetOptions> at offset %d: %s", offset, err), nil
					}
				}
			case "Table":
				if 2 == level {
					_ = xls.iteratorDecoder.Skip()
				}
			default:
				_ = xls.iteratorDecoder.Skip()
			}
		}
	}
	_ = xls.SetSheetId(xls.sheetSelected)
	return nil, xls
}

func (sheet *xmlTableSheetInfo) GetName() string {
	return sheet.Name
}

func (sheet *xmlTableSheetInfo) GetHideLevel() TSheetHideLevel {
	return sheet.HideLevel
}

func (xls *xmlHandle) Close() error {
	return xls.iteratorStreamSource.Close()
}

func (sheet *xmlHandle) FormatterAvailable() bool {
	return false
}

func (xls *xmlHandle) SetI18n(string) error {
	_, _ = os.Stderr.WriteString("WARNING! Formatter is unavailable for XLS format [1]!\n")
	return nil
}

func (xls *xmlHandle) Formatter() IExcelFormatter {
	_, _ = os.Stderr.WriteString("WARNING! Formatter is unavailable for XLS format [2]!\n")
	return newExcelFormatter("en")
}

func (xls *xmlHandle) GetSheets() []ITableSheetInfo {
	res := make([]ITableSheetInfo, len(xls.sheets))
	for i, sheet := range xls.sheets {
		res[i] = sheet
	}
	return res
}

func (xls *xmlHandle) GetCurrentSheetId() int {
	return xls.iteratorSheetId
}

func (xls *xmlHandle) SetSheetId(id int) error {
	xls.iteratorLastError = nil
	xls.iteratorCapacity = 0
	xls.iteratorRowNum = 0
	xls.iteratorScannedData = []string{}
	xls.iteratorXMLSegment = iteratorRXSegmentRoot
	if id < 0 || id > len(xls.sheets) {
		return fmt.Errorf("sheet #%d not found", id)
	}
	xls.iteratorSheetId = id
	xls.iteratorDecoder = nil
	return nil
}

func (xls *xmlHandle) GetLastScanError() error {
	return xls.iteratorLastError
}

func (xls *xmlHandle) Scan() (err error) {
	// if row we have scanned is not next to previously returned, just increase "previouslyReturned" counter and imply empty row
	if xls.iteratorScannedRowNum > xls.iteratorRowNum {
		xls.iteratorRowNum++
	} else {
		err = xls.scanInternal()
		if nil == err {
			xls.iteratorRowNum++
		}
	}
	xls.iteratorLastError = err
	return err
}

func (xlsx *xmlHandle) GetScanned() []string {
	if xlsx.iteratorScannedRowNum > xlsx.iteratorRowNum {
		return []string{}
	}
	return xlsx.iteratorScannedData
}

func (xls *xmlHandle) requireScanStream() error {
	if nil == xls.iteratorDecoder {
		xls.iteratorDecoderInitialOffset = xls.sheets[xls.iteratorSheetId].start
		_, err := xls.iteratorStreamXML.Seek(xls.iteratorDecoderInitialOffset, io.SeekStart)
		if nil != err {
			return fmt.Errorf("seek [%d] failed, some file contents are missing", xls.sheets[xls.iteratorSheetId].start)
		}
		xls.iteratorDecoder = xml.NewDecoder(xls.iteratorStreamXML)
	}
	return nil
}

func (xls *xmlHandle) scanInternal() error {
	err := xls.requireScanStream()
	if nil != err {
		return err
	}
	//var level byte = 0    // 0=./ 1=./Worksheet 2=./Worksheet/Table  3=./Worksheet/Table/Row  4=./Worksheet/Table/Row/Cell*
	xls.iteratorScannedData = make([]string, 0, xls.iteratorCapacity)
	rowIsParsed := false
	for !rowIsParsed {
		var tokenErr error
		var tok xml.Token
		offset := xls.iteratorDecoderInitialOffset + xls.iteratorDecoder.InputOffset()
		if offset > xls.sheets[xls.iteratorSheetId].stop {
			// do not iterate out of sheet offset bounds
			tokenErr = io.EOF
		} else {
			tok, tokenErr = xls.iteratorDecoder.Token()
			if tokenErr == nil && offset > xls.sheets[xls.iteratorSheetId].stop {
				// do not iterate out of sheet offset bounds
				tokenErr = io.EOF
			}
			if len(xls.iteratorScannedData) > xls.iteratorCapacity {
				xls.iteratorCapacity = len(xls.iteratorScannedData)
			}
		}
		if tokenErr != nil || tok == nil {
			_ = xls.SetSheetId(xls.iteratorSheetId)
			if io.EOF == tokenErr {
				return tokenErr
			}
			return fmt.Errorf("xml token read error in at pos %d: %s", offset, tokenErr.Error())
		}
		switch tok := tok.(type) {
		case xml.EndElement:
			switch tok.Name.Local {
			case "Worksheet":
				if iteratorRXSegmentW == xls.iteratorXMLSegment {
					xls.iteratorXMLSegment = iteratorRXSegmentRoot
				}
			case "Table":
				if iteratorRXSegmentWT == xls.iteratorXMLSegment {
					xls.iteratorXMLSegment = iteratorRXSegmentW
				}
			case "Row":
				rowIsParsed = true
				if iteratorRXSegmentWTR == xls.iteratorXMLSegment {
					xls.iteratorXMLSegment = iteratorRXSegmentWT
				}
			case "Cell":
				if iteratorRXSegmentWTRC == xls.iteratorXMLSegment {
					xls.iteratorXMLSegment = iteratorRXSegmentWTR
				}
			}
		case xml.StartElement:
			switch tok.Name.Local {
			case "Worksheet":
				if iteratorRXSegmentRoot == xls.iteratorXMLSegment {
					xls.iteratorXMLSegment = iteratorRXSegmentW
				}
			case "Table":
				if iteratorRXSegmentW == xls.iteratorXMLSegment {
					xls.iteratorXMLSegment = iteratorRXSegmentWT
				}
			case "Row":
				if iteratorRXSegmentWT == xls.iteratorXMLSegment {
					xls.iteratorXMLSegment = iteratorRXSegmentWTR
					xls.iteratorScannedData = make([]string, 0, xls.iteratorCapacity)
					err, currentRowNumStr := findXmlTokenAttrValue(&tok, "Index")
					if nil == err {
						attrNum, err := strconv.Atoi(currentRowNumStr)
						if nil != err {
							return fmt.Errorf("cannot parse <Row> Index attr at offset %d", offset)
						}
						if attrNum <= xls.iteratorScannedRowNum {
							return fmt.Errorf("<Cell Index> collision detected at offset %d", offset)
						}
						xls.iteratorScannedRowNum = attrNum
					} else {
						xls.iteratorScannedRowNum++
					}
				}
			case "Cell":
				if iteratorRXSegmentWTR == xls.iteratorXMLSegment {
					// @todo: implement ss:MergeAcross
					currentColumnNum := 0 // 1-based
					err, colNumStr := findXmlTokenAttrValue(&tok, "Index")
					mergeNum := 0
					if nil == err {
						currentColumnNum, err = strconv.Atoi(colNumStr)
						if nil != err {
							return fmt.Errorf("cannot parse <Row>#%d<Cell> Index attr at offset %d", xls.iteratorScannedRowNum, offset)
						}
					} else {
						currentColumnNum = len(xls.iteratorScannedData) + 1
					}
					err, mergeStr := findXmlTokenAttrValue(&tok, "MergeAcross")
					if nil == err {
						mergeNum, err = strconv.Atoi(mergeStr)
						if nil != err {
							return fmt.Errorf("cannot parse <Row>#%d<Cell>#%d MergeAcross attr at offset %d", xls.iteratorScannedRowNum, currentColumnNum, offset)
						}
					}
					cell := &rawxmlCell{}
					err = xls.iteratorDecoder.DecodeElement(cell, &tok)
					if len(xls.iteratorScannedData) > currentColumnNum-1 {
						return fmt.Errorf("cell index should be greater than previous cell's one <Row>#%d<Cell>#%d at offset %d", xls.iteratorScannedRowNum, currentColumnNum, offset)
					} else {
						for len(xls.iteratorScannedData) < currentColumnNum-1 {
							xls.iteratorScannedData = append(xls.iteratorScannedData, "")
						}
						xls.iteratorScannedData = append(xls.iteratorScannedData, cell.Data)
					}
					for i := 0; i < mergeNum; i++ {
						xls.iteratorScannedData = append(xls.iteratorScannedData, "")
					}
				}
			default:
				_ = xls.iteratorDecoder.Skip()
			}
		}
	}
	return nil
}
