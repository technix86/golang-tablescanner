package tablescanner

// @todo: implement xls support
// @todo: implement csv support
// @todo: implement excel-xml support
// @todo: implement html support
// @todo: gen tests
// @todo someday: merged cell behaviour (mode=[none,showClone,showRef],directions:[row,cell,table])
//                NB keep the mind on existing row-skip-behaviour

import (
	"io"
)

type TSheetHideLevel byte

const (
	TableSheetVisible    TSheetHideLevel = 0
	TableSheetHidden     TSheetHideLevel = 1
	TableSheetVeryHidden TSheetHideLevel = 2
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
