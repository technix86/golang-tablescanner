package tablescanner

import (
	"archive/zip"
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

type ITableDocumentScanner interface {
	io.Closer
	GetSheets() []TableSheetInfo
	GetCurrentSheetId() int
	SetSheetId(id int) error
	Scan() error
	GetScanned() []string
}

func NewXLSXStream(fileName string) (error, ITableDocumentScanner) {
	var err error
	xlsx := &xlsxStream{zFileName: fileName}
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
	err = xlsx.readSharedStrings("xl/sharedStrings.xml")
	if err != nil {
		return err, nil
	}
	err = xlsx.readStyles("xl/styles.xml")
	if err != nil {
		return err, nil
	}
	err = xlsx.readWorkbook("xl/workbook.xml")
	if err != nil {
		return err, nil
	}
	return nil, xlsx
}
