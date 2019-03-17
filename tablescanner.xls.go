package tablescanner

import (
	"fmt"
	exls "github.com/extrame/xls"
	"io"
	"os"
)

type xlsTableSheetInfo struct {
	Name  string
	sheet *exls.WorkSheet
}

type xlsHandle struct {
	formatter           excelFormatter
	sheets              []*xlsTableSheetInfo
	sheetSelected       int      // default-opening sheet id
	iteratorLastError   error    // error which caused last Scan() failed
	iteratorScannedData []string // current row-iterating row data
	iteratorRowNum      int      // row number that Scan() implies (starting with 1)
	iteratorSheetId     int      // current row-iterating sheet id
	closer              io.Closer
	workbook            *exls.WorkBook
}

func newXLSStream(fileName string) (error, ITableDocumentScanner) {
	var err error
	xls := &xlsHandle{}
	xls.workbook, xls.closer, err = exls.OpenWithCloser(fileName, "utf-8")
	err = xls.SetI18n("en")
	if err == nil {
		return err, nil
	}
	numSheets := xls.workbook.NumSheets()
	xls.sheets = make([]*xlsTableSheetInfo, numSheets)
	for i := 0; i < numSheets; i++ {
		xsheet := xls.workbook.GetSheet(i)
		xls.sheets[i] = &xlsTableSheetInfo{Name: xsheet.Name, sheet: xsheet}
	}
	panic("@todo: search for current sheet selected")
	return nil, xls
}

func (sheet *xlsTableSheetInfo) GetName() string {
	return sheet.Name
}

func (sheet *xlsTableSheetInfo) GetHideLevel() TSheetHideLevel {
	return TableSheetVisible
}

func (xls *xlsHandle) Close() error {
	return xls.closer.Close()
}

func (xls *xlsHandle) SetI18n(string) error {
	_, _ = os.Stderr.WriteString("WARNING! Formatter is unavailable for XLS format!")
	return nil
}

func (xls *xlsHandle) Formatter() IExcelFormatter {
	_, _ = os.Stderr.WriteString("WARNING! Formatter is unavailable for XLS format!")
	return newExcelFormatter()
}

func (xls *xlsHandle) GetSheets() []ITableSheetInfo {
	res := make([]ITableSheetInfo, len(xls.sheets))
	for i, sheet := range xls.sheets {
		res[i] = sheet
	}
	return res
}

func (xls *xlsHandle) GetCurrentSheetId() int {
	return xls.iteratorSheetId
}

func (xls *xlsHandle) SetSheetId(id int) error {
	xls.iteratorLastError = nil
	xls.iteratorRowNum = 0
	xls.iteratorScannedData = []string{}
	if id < 0 || id > len(xls.sheets) {
		return fmt.Errorf("sheet #%d not found", id)
	}
	xls.iteratorSheetId = id
	return nil
}

func (xls *xlsHandle) GetLastScanError() error {
	return xls.iteratorLastError
}

func (xls *xlsHandle) Scan() error {
	xls.iteratorRowNum++
	xls.iteratorLastError = xls.scanInternal()
	return xls.iteratorLastError
}

func (xls *xlsHandle) GetScanned() []string {
	return xls.iteratorScannedData
}

func (xls *xlsHandle) scanInternal() error {
	ROWSTARTING := 0
	COLSTARTING := 1
	if xls.iteratorSheetId < 0 || xls.iteratorSheetId > len(xls.sheets) {
		return fmt.Errorf("sheet #%d not found", xls.iteratorSheetId)
	}
	if xls.iteratorRowNum < 0 {
		panic("xls.iteratorRowNum must be >= 1")
	}
	if xls.iteratorRowNum > int(xls.sheets[xls.iteratorSheetId].sheet.MaxRow) {
		return io.EOF
	}
	row := xls.sheets[xls.iteratorSheetId].sheet.Row(xls.iteratorRowNum - 1 + ROWSTARTING)
	if nil == row {
		xls.iteratorScannedData = make([]string, 0)
		return nil

	}
	colFirst, colLast := row.FirstCol(), row.LastCol()
	if colLast < colFirst {
		return fmt.Errorf("invalid data for row #%d, FirstCol()=%d > LastCol()=%d", xls.iteratorRowNum, colFirst, colLast)
	}
	colFirst -= COLSTARTING
	colLast -= COLSTARTING
	xls.iteratorScannedData = make([]string, colLast+1, colLast+1)
	for i := colFirst; i < colLast; i++ {
		xls.iteratorScannedData[i] = row.Col(i)
	}
	return nil
}
