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

const (
	TableSheetVisible    = 0
	TableSheetHidden     = 1
	TableSheetVeryHidden = 2
)

type TableSheetInfo struct {
	Name      string
	HideLevel byte
	path      string
	rId       string
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
	GetSheets() []TableSheetInfo
	GetCurrentSheetId() int
	SetSheetId(id int) error
	Scan() error
	GetLastScanError() error
	GetScanned() []string
}

func NewXLSXStream(fileName string) (error, ITableDocumentScanner) {
	return newXLSXStream(fileName)
}
