package tablescanner

import (
	"io"
)

// @todo someday: merged cell behaviour (mode=[none,showClone,showRef],directions:[row,cell,table]), NB about existing row-skip-behaviour
// @todo: gen tests

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
	FormatValue(cellValue string, cellType string, fullFormat *parsedNumberFormat) (string, error)
}

type ITableDocumentScanner interface {
	io.Closer
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
