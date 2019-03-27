package tablescanner

// @todo: detect xlsx additionally by required subfile /xl/workbook.xml to prevent false positive on common ZIPs
// @todo: implement csv support
// @todo: implement excel-xml support
// @todo: implement html support
// @todo: gen tests

import (
	"bytes"
	"fmt"
	"io"
	"os"
)

type TSheetHideLevel byte

const (
	TableSheetVisible    TSheetHideLevel = 0
	TableSheetHidden     TSheetHideLevel = 1
	TableSheetVeryHidden TSheetHideLevel = 2
)

type TExcelWorkbookType byte

const (
	TypeExcelWorkbookUnknown TExcelWorkbookType = 0
	TypeExcelWorkbookXLSX    TExcelWorkbookType = 1
	TypeExcelWorkbookXLS     TExcelWorkbookType = 2
	TypeExcelWorkbookXML     TExcelWorkbookType = 3
)

type ITableSheetInfo interface {
	GetName() string
	GetHideLevel() TSheetHideLevel
}

type IExcelFormatter interface {
	DisableFormatting()
	EnableFormatting()
	AllowScientific()
	DenyScientific()
	SetDateFixedFormat(value string)
	SetDecimalSeparator(value string)
	SetThousandSeparator(value string)
	SetTrimOn()
	SetTrimOff()
	FormatValue(cellValue string, cellType string, fullFormat *parsedNumberFormat) (string, error)
}

type ITableDocumentScanner interface {
	io.Closer
	SetI18n(string) error
	Formatter() IExcelFormatter
	GetSheets() []ITableSheetInfo
	GetCurrentSheetId() int
	SetSheetId(id int) error
	Scan() error
	GetLastScanError() error
	GetScanned() []string
}

func NewXLSXStream(fileName string) (error, ITableDocumentScanner) {
	return newXLSXStream(fileName)
}

func NewXLSStream(fileName string) (error, ITableDocumentScanner) {
	return newXLSStream(fileName)
}

func NewTableStream(fileName string) (error, ITableDocumentScanner) {
	err, excelType := DetectExcelContentType(fileName)
	if nil != err {
		return err, nil
	}
	switch excelType {
	case TypeExcelWorkbookXLSX:
		return NewXLSXStream(fileName)
	case TypeExcelWorkbookXLS:
		return NewXLSStream(fileName)
	}
	return fmt.Errorf("file %s has unsupported format", fileName), nil
}

func DetectExcelContentType(fileName string) (error, TExcelWorkbookType) {
	signatureXLSX := []byte("\x50\x4B\x03\x04\x14")
	signatureXLS := []byte("\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1")
	signatureXML := []byte("\xFF\xFE\x3C\x00\x3F\x00\x78\x00")
	file, err := os.Open(fileName)
	if nil != err {
		return err, TypeExcelWorkbookUnknown
	}
	defer nowarnCloseCloser(file)
	signature := make([]byte, 8)
	bytesRead, err := file.Read(signature)
	if err != nil {
		return fmt.Errorf("cannot detect content type of file %s: %s", fileName, err), TypeExcelWorkbookUnknown
	}
	if 8 != bytesRead {
		// @todo: does Read() return error if bytesRead<8 ?
		return nil, TypeExcelWorkbookUnknown
	}
	if bytes.Equal(signatureXLSX, signature[0:len(signatureXLSX)]) {
		return nil, TypeExcelWorkbookXLSX
	}
	if bytes.Equal(signatureXLS, signature[0:len(signatureXLS)]) {
		return nil, TypeExcelWorkbookXLS
	}
	if bytes.Equal(signatureXML, signature[0:len(signatureXML)]) {
		return nil, TypeExcelWorkbookXML
	}
	return nil, TypeExcelWorkbookUnknown
}

func nowarnCloseCloser(rc io.Closer) {
	_ = rc.Close()
}
