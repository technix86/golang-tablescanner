package tablescanner

// @todo: detect xlsx additionally by required subfile /xl/workbook.xml to prevent false positive on common ZIPs
// @todo: implement csv support
// @todo: implement excel-xml support
// @todo: implement html support
// @todo: gen tests

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"os"
	"unicode/utf16"
	"unicode/utf8"
)

type TSheetHideLevel byte

const (
	TableSheetVisible    TSheetHideLevel = 0
	TableSheetHidden     TSheetHideLevel = 1
	TableSheetVeryHidden TSheetHideLevel = 2
)

type TTextEnconding byte

const (
	EncodingUnknown TTextEnconding = 0
	EncodingUTF8    TTextEnconding = 1
	EncodingUTF16BE TTextEnconding = 2
	EncodingUTF16LE TTextEnconding = 3
)

var signatureBOMUTF8 = []byte("\xEF\xBB\xBF")
var signatureBOMUTF16BE = []byte("\xFE\xFF")
var signatureBOMUTF16LE = []byte("\xFF\xFE")

type TExcelWorkbookType byte

const (
	TypeExcelWorkbookUnknown    TExcelWorkbookType = 0
	TypeExcelWorkbookXLSX       TExcelWorkbookType = 1
	TypeExcelWorkbookXLS        TExcelWorkbookType = 2
	TypeExcelWorkbookXML        TExcelWorkbookType = 3
	TypeExcelWorkbookSingleHTML TExcelWorkbookType = 4
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
	FormatterAvailable() bool
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
	err, excelType, textEncoding,bomPresent := DetectExcelContentType(fileName)
	if nil != err {
		return err, nil
	}
	switch excelType {
	case TypeExcelWorkbookXLSX:
		return NewXLSXStream(fileName)
	case TypeExcelWorkbookXLS:
		return NewXLSStream(fileName)
	case TypeExcelWorkbookXML:
		return newXMLStream(fileName,textEncoding,bomPresent)
	}
	return fmt.Errorf("file %s has unsupported format", fileName), nil
}

func DetectExcelContentType(fileName string) (err error, bookType TExcelWorkbookType, textEncoding TTextEnconding, BOMPresent []byte) {
	signatureXLSX := []byte("\x50\x4B\x03\x04\x14")
	signatureXLS := []byte("\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1")
	signatureXML := []byte("<?xml")
	signatureHTML := []byte("<html")
	//
	bookType = TypeExcelWorkbookUnknown
	textEncoding = EncodingUnknown
	file, err := os.Open(fileName)
	if nil != err {
		bookType = TypeExcelWorkbookUnknown
		return
	}
	defer nowarnCloseCloser(file)
	signature := make([]byte, 64)
	_, err = file.Read(signature)
	if err != nil {
		err = fmt.Errorf("cannot detect content type of file %s: %s", fileName, err)
		return
	}
	if len(signature) >= len(signatureXLSX) && bytes.Equal(signatureXLSX, signature[0:len(signatureXLSX)]) {
		bookType = TypeExcelWorkbookXLSX
		textEncoding = EncodingUTF8
		return
	}
	if len(signature) >= len(signatureXLS) && bytes.Equal(signatureXLS, signature[0:len(signatureXLS)]) {
		bookType = TypeExcelWorkbookXLS
		textEncoding = EncodingUTF8
		return
	}
	// text-based formats allowed below this point only
	if len(signature) >= len(signatureBOMUTF8) && bytes.Equal(signatureBOMUTF8, signature[0:len(signatureBOMUTF8)]) {
		BOMPresent = signatureBOMUTF8
		textEncoding = EncodingUTF8
		signature = signature[len(signatureBOMUTF8):]
	} else if len(signature) >= len(signatureBOMUTF16BE) && bytes.Equal(signatureBOMUTF16BE, signature[0:len(signatureBOMUTF16BE)]) {
		BOMPresent = signatureBOMUTF16BE
		textEncoding = EncodingUTF16BE
		signature = signature[len(signatureBOMUTF16BE):]
		signature = UTF16BytesToUTF8Bytes(signature, binary.BigEndian)
	} else if len(signature) >= len(signatureBOMUTF16LE) && bytes.Equal(signatureBOMUTF16LE, signature[0:len(signatureBOMUTF16LE)]) {
		BOMPresent = signatureBOMUTF16LE
		textEncoding = EncodingUTF16LE
		signature = signature[len(signatureBOMUTF16LE):]
		signature = UTF16BytesToUTF8Bytes(signature, binary.LittleEndian)
	}
	if len(signature) >= len(signatureXML) && bytes.Equal(signatureXML, signature[0:len(signatureXML)]) {
		return nil, TypeExcelWorkbookXML, textEncoding, BOMPresent
	}
	if len(signature) >= len(signatureHTML) && bytes.Equal(signatureHTML, signature[0:len(signatureHTML)]) {
		return nil, TypeExcelWorkbookSingleHTML, textEncoding, BOMPresent
	}
	return nil, TypeExcelWorkbookUnknown, EncodingUnknown, BOMPresent
}

func UTF16BytesToUTF8Bytes(b []byte, o binary.ByteOrder) []byte {
	utf := make([]uint16, (len(b)+1)/2)
	for i := 0; i+1 < len(b); i += 2 {
		utf[i/2] = o.Uint16(b[i:])
	}
	if len(b)/2 < len(utf) {
		utf[len(utf)-1] = utf8.RuneError
	}
	return []byte(string(utf16.Decode(utf)))
}

func nowarnCloseCloser(rc io.Closer) {
	_ = rc.Close()
}
