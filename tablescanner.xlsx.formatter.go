package tablescanner

import (
	"errors"
	"fmt"
	"math"
	"strconv"
	"strings"
	"time"
)

const (
	MJD_0                    = float64(2400000.5)
	nanosInADay              = float64((24 * time.Hour) / time.Nanosecond)
	strCellTypeError         = "e"
	strCellTypeString        = "s"
	strCellTypeInline        = "inlineStr"
	strCellTypeBool          = "b"
	strCellTypeStringFormula = "str"
	strCellTypeDate          = "d"
	strCellTypeNumeric       = "n"
	strCellTypeNumericAlt    = ""
	maxNonScientificNumber   = 1e11
	minNonScientificNumber   = 1e-9
)

type excelFormatter struct {
	i18n              *tI18n
	discardFormatting bool
	allowScientific   bool
	dateFixedFormat   string
	decimalSeparator  string
	thousandSeparator string
	trim              bool
	date1904          bool
}

type parsedNumberFormat struct {
	numFmt                        string
	isTimeFormat                  bool
	negativeFormatExpectsPositive bool
	positiveFormat                *formatOptions
	negativeFormat                *formatOptions
	zeroFormat                    *formatOptions
	textFormat                    *formatOptions
	parseEncounteredError         *error
}

type formatOptions struct {
	isTimeFormat        bool
	showPercent         bool
	fullFormatString    string
	reducedFormatString string
	prefix              string
	suffix              string
}

var excel1900Epoc = time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)
var excel1904Epoc = time.Date(1904, time.January, 1, 0, 0, 0, 0, time.UTC)

var timeFormatCharacters = []string{"m", "d", "yy", "h", "m", "AM/PM", "A/P", "am/pm", "a/p", "r", "g", "e", "b1", "b2", "[hh]", "[h]", "[mm]", "[m]",
	"s.0000", "s.000", "s.00", "s.0", "s", "[ss].0000", "[ss].000", "[ss].00", "[ss].0", "[ss]", "[s].0000", "[s].000", "[s].00", "[s].0", "[s]", "г", "г."}

var formattingCharacters = []string{"0/", "#/", "?/", "E-", "E+", "e-", "e+", "0", "#", "?", ".", ",", "@", "*"}

var fallbackErrorFormat = &formatOptions{
	fullFormatString:    "general",
	reducedFormatString: "general",
}

func newExcelFormatter() *excelFormatter {
	return &excelFormatter{}
}
func (formatter *excelFormatter) setDate1904(date1904 bool) {
	formatter.date1904 = date1904
}
func (formatter *excelFormatter) setI18n(i18n *tI18n) {
	formatter.i18n = i18n
	formatter.decimalSeparator = i18n.decimalSeparator
	formatter.thousandSeparator = i18n.thousandSeparator
}

func (formatter *excelFormatter) DisableFormatting() {
	formatter.discardFormatting = true
}
func (formatter *excelFormatter) EnableFormatting() {
	formatter.discardFormatting = false
}
func (formatter *excelFormatter) AllowScientific() {
	formatter.allowScientific = true
}
func (formatter *excelFormatter) DenyScientific() {
	formatter.allowScientific = false
}
func (formatter *excelFormatter) SetTrimOn() {
	formatter.trim = true
}
func (formatter *excelFormatter) SetTrimOff() {
	formatter.trim = false
}
func (formatter *excelFormatter) SetDateFixedFormat(value string) {
	formatter.dateFixedFormat = value
}

func (formatter *excelFormatter) SetDecimalSeparator(value string) {
	if value == "" {
		value = formatter.i18n.decimalSeparator
	}
	formatter.decimalSeparator = value
}

func parseNumFmt(numFmt string) *parsedNumberFormat {
	if "" == numFmt {
		numFmt = "general"
	}
	parsedNumFmt := &parsedNumberFormat{
		numFmt: numFmt,
	}
	if isTimeFormat(numFmt) {
		// Time formats cannot have multiple groups separated by semicolons, there is only one format.
		// Strings are unaffected by the time format.
		parsedNumFmt.isTimeFormat = true
		parsedNumFmt.textFormat, _ = parseNumberFormatSection("general")
		return parsedNumFmt
	}

	var fmtOptions []*formatOptions
	formats, err := splitFormatOnSemicolon(numFmt)
	if err == nil {
		for _, formatSection := range formats {
			parsedFormat, err := parseNumberFormatSection(formatSection)
			if err != nil {
				// If an invalid number section is found, fall back to general
				parsedFormat = fallbackErrorFormat
				parsedNumFmt.parseEncounteredError = &err
			}
			fmtOptions = append(fmtOptions, parsedFormat)
		}
	} else {
		fmtOptions = append(fmtOptions, fallbackErrorFormat)
		parsedNumFmt.parseEncounteredError = &err
	}
	if len(fmtOptions) > 4 {
		fmtOptions = []*formatOptions{fallbackErrorFormat}
		err = errors.New("invalid number format, too many format sections")
		parsedNumFmt.parseEncounteredError = &err
	}

	if len(fmtOptions) == 1 {
		// If there is only one option, it is used for all
		parsedNumFmt.positiveFormat = fmtOptions[0]
		parsedNumFmt.negativeFormat = fmtOptions[0]
		parsedNumFmt.zeroFormat = fmtOptions[0]
		if strings.Contains(fmtOptions[0].fullFormatString, "@") {
			parsedNumFmt.textFormat = fmtOptions[0]
		} else {
			parsedNumFmt.textFormat, _ = parseNumberFormatSection("general")
		}
	} else if len(fmtOptions) == 2 {
		// If there are two formats, the first is used for positive and zeros, the second gets used as a negative format,
		// and strings are not formatted.
		// When negative numbers now have their own format, they should become positive before having the format applied.
		// The format will contain a negative sign if it is desired, but they may be colored red or wrapped in
		// parenthesis instead.
		parsedNumFmt.negativeFormatExpectsPositive = true
		parsedNumFmt.positiveFormat = fmtOptions[0]
		parsedNumFmt.negativeFormat = fmtOptions[1]
		parsedNumFmt.zeroFormat = fmtOptions[0]
		parsedNumFmt.textFormat, _ = parseNumberFormatSection("general")
	} else if len(fmtOptions) == 3 {
		// If there are three formats, the first is used for positive, the second gets used as a negative format,
		// the third is for negative, and strings are not formatted.
		parsedNumFmt.negativeFormatExpectsPositive = true
		parsedNumFmt.positiveFormat = fmtOptions[0]
		parsedNumFmt.negativeFormat = fmtOptions[1]
		parsedNumFmt.zeroFormat = fmtOptions[2]
		parsedNumFmt.textFormat, _ = parseNumberFormatSection("general")
	} else {
		// With four options, the first is positive, the second is negative, the third is zero, and the fourth is strings
		// Negative numbers should be still become positive before having the negative formatting applied.
		parsedNumFmt.negativeFormatExpectsPositive = true
		parsedNumFmt.positiveFormat = fmtOptions[0]
		parsedNumFmt.negativeFormat = fmtOptions[1]
		parsedNumFmt.zeroFormat = fmtOptions[2]
		parsedNumFmt.textFormat = fmtOptions[3]
	}
	return parsedNumFmt
}

func parseLiterals(format string) (string, string, bool, error) {
	var prefix string
	showPercent := false
	for i := 0; i < len(format); i++ {
		curReducedFormat := format[i:]
		switch curReducedFormat[0] {
		case '\\':
			// If there is a slash, skip the next character, and add it to the prefix
			if len(curReducedFormat) > 1 {
				i++
				prefix += curReducedFormat[1:2]
			}
		case '_':
			// If there is an underscore, skip the next character, but don't add it to the prefix
			if len(curReducedFormat) > 1 {
				i++
			}
		case '*':
			// Asterisks are used to repeat the next character to fill the full cell width.
			// There isn't really a cell size in this context, so this will be ignored.
		case '"':
			// If there is a quote skip to the next quote, and add the quoted characters to the prefix
			endQuoteIndex := strings.Index(curReducedFormat[1:], "\"")
			if endQuoteIndex == -1 {
				return "", "", false, errors.New("invalid formatting code, unmatched double quote")
			}
			prefix = prefix + curReducedFormat[1:endQuoteIndex+1]
			i += endQuoteIndex + 1
		case '%':
			showPercent = true
			prefix += "%"
		case '[':
			// Brackets can be currency annotations (e.g. [$$-409])
			// color formats (e.g. [color1] through [color56], as well as [red] etc.)
			// conditionals (e.g. [>100], the valid conditionals are =, >, <, >=, <=, <>)
			bracketIndex := strings.Index(curReducedFormat, "]")
			if bracketIndex == -1 {
				return "", "", false, errors.New("invalid formatting code, invalid brackets")
			}
			// Currencies in Excel are annotated with this format: [$<Currency String>-<Language Info>]
			// Currency String is something like $, ¥, €, or £
			// Language Info is three hexadecimal characters
			if len(curReducedFormat) > 2 && curReducedFormat[1] == '$' {
				dashIndex := strings.Index(curReducedFormat, "-")
				if dashIndex != -1 && dashIndex < bracketIndex {
					// Get the currency symbol, and skip to the end of the currency format
					prefix += curReducedFormat[2:dashIndex]
				} else {
					return "", "", false, errors.New("invalid formatting code, invalid currency annotation")
				}
			}
			i += bracketIndex
		case '$', '-', '+', '/', '(', ')', ':', '!', '^', '&', '\'', '~', '{', '}', '<', '>', '=', ' ':
			// These symbols are allowed to be used as literal without escaping
			prefix += curReducedFormat[0:1]
		default:
			for _, special := range formattingCharacters {
				if strings.HasPrefix(curReducedFormat, special) {
					// This means we found the start of the actual number formatting portion, and should return.
					return prefix, format[i:], showPercent, nil
				}
			}
			// Symbols that don't have meaning and aren't in the exempt literal characters and are not escaped.
			return "", "", false, errors.New("invalid formatting code: unsupported or unescaped characters")
		}
	}
	return prefix, "", showPercent, nil
}

func parseNumberFormatSection(fullFormat string) (*formatOptions, error) {
	reducedFormat := strings.TrimSpace(fullFormat)

	// general is the only format that does not use the normal format symbols notations
	if reducedFormat == "" || strings.ToLower(reducedFormat) == "general" {
		return &formatOptions{
			fullFormatString:    "general",
			reducedFormatString: "general",
		}, nil
	}

	prefix, reducedFormat, showPercent1, err := parseLiterals(reducedFormat)
	if err != nil {
		return nil, err
	}

	reducedFormat, suffixFormat := splitFormatAndSuffixFormat(reducedFormat)

	suffix, remaining, showPercent2, err := parseLiterals(suffixFormat)
	if err != nil {
		return nil, err
	}
	if len(remaining) > 0 {
		// This paradigm of codes consisting of literals, number formats, then more literals is not always correct, they can
		// actually be intertwined. Though 99% of the time number formats will not do this.
		// Excel uses this format string for Social Security Numbers: 000\-00\-0000
		// and this for US phone numbers: [<=9999999]###\-####;\(###\)\ ###\-####
		return nil, errors.New("invalid or unsupported format string")
	}

	if -1 != strings.IndexRune(reducedFormat, ';') {
		reducedFormat = reducedFormat[0:strings.IndexRune(reducedFormat, ';')]
	}
	return &formatOptions{
		fullFormatString:    fullFormat,
		isTimeFormat:        false,
		reducedFormatString: reducedFormat,
		prefix:              prefix,
		suffix:              suffix,
		showPercent:         showPercent1 || showPercent2,
	}, nil
}
func splitFormatAndSuffixFormat(format string) (string, string) {
	var i int
	for ; i < len(format); i++ {
		curReducedFormat := format[i:]
		var found bool
		for _, special := range formattingCharacters {
			if strings.HasPrefix(curReducedFormat, special) {
				// Skip ahead if the special character was longer than length 1
				i += len(special) - 1
				found = true
				break
			}
		}
		if !found {
			break
		}
	}
	suffixFormat := format[i:]
	format = format[:i]
	return format, suffixFormat
}

func splitFormatOnSemicolon(format string) ([]string, error) {
	var formats []string
	prevIndex := 0
	for i := 0; i < len(format); i++ {
		if format[i] == ';' {
			formats = append(formats, format[prevIndex:i])
			prevIndex = i + 1
		} else if format[i] == '\\' {
			i++
		} else if format[i] == '"' {
			endQuoteIndex := strings.Index(format[i+1:], "\"")
			if endQuoteIndex == -1 {
				// This is an invalid format string, fall back to general
				return nil, errors.New("invalid format string, unmatched double quote")
			}
			i += endQuoteIndex + 1
		}
	}
	return append(formats, format[prevIndex:]), nil
}
func (formatter *excelFormatter) FormatValue(cellValue string, cellType string, fullFormat *parsedNumberFormat) (string, error) {
	result, err := formatter.internalFormatValue(cellValue, cellType, fullFormat)
	if nil == err && formatter.trim {
		result = strings.TrimSpace(result)
	}
	return result, err
}

func (formatter *excelFormatter) internalFormatValue(cellValue string, cellType string, fullFormat *parsedNumberFormat) (string, error) {
	if formatter.discardFormatting {
		return cellValue, nil
	}
	switch cellType {
	case strCellTypeError:
		// The error type is what XLSX uses in error cases such as when formulas are invalid.
		// There will be text in the cell's value that can be shown, something ugly like #NAME? or #######
		return cellValue, nil
	case strCellTypeBool:
		if cellValue == "0" {
			return "FALSE", nil
		} else if cellValue == "1" {
			return "TRUE", nil
		} else {
			return cellValue, errors.New("invalid value in bool cell")
		}
	case strCellTypeString:
		fallthrough
	case strCellTypeInline:
		fallthrough
	case strCellTypeStringFormula:
		textFormat := fullFormat.textFormat
		// This switch statement is only for String formats
		switch textFormat.reducedFormatString {
		case "general":
			return cellValue, nil
		case "@": // String is "@"
			return textFormat.prefix + cellValue + textFormat.suffix, nil
		case "":
			// If cell is not "General" and there is not an "@" symbol in the format, then the cell's value is not
			// used when determining what to display. It would be completely legal to have a format of "Error"
			// for strings, and all values that are not numbers would show up as "Error". In that case, this code would
			// have a prefix of "Error" and a reduced format string of "" (empty string).
			return textFormat.prefix + textFormat.suffix, nil
		default:
			return cellValue, errors.New("invalid or unsupported format, unsupported string format")
		}
	case strCellTypeDate:
		// These are dates that are stored in date format instead of being stored as numbers with a format to turn them
		// into a date string.
		return cellValue, nil
	case strCellTypeNumeric:
		fallthrough
	case strCellTypeNumericAlt:
		return formatter.formatNumericCell(cellValue, fullFormat)
	default:
		return cellValue, errors.New("unknown cell type")
	}
}

func (formatter *excelFormatter) formatNumericCell(cellValue string, fullFormat *parsedNumberFormat) (string, error) {
	rawValue := strings.TrimSpace(cellValue)
	// If there wasn't a value in the cell, it shouldn't have been marked as Numeric.
	// It's better to support this case though.
	if rawValue == "" {
		return "", nil
	}

	if fullFormat.isTimeFormat {
		return formatter.parseTime(rawValue, fullFormat)
	}
	var numberFormat *formatOptions
	floatVal, floatErr := strconv.ParseFloat(rawValue, 64)
	if floatErr != nil {
		return rawValue, floatErr
	}
	// Choose the correct format. There can be different formats for positive, negative, and zero numbers.
	// Excel only uses the zero format if the value is literally zero, even if the number is so small that it shows
	// up as "0" when the positive format is used.
	if floatVal > 0 {
		numberFormat = fullFormat.positiveFormat
	} else if floatVal < 0 {
		// If format string specified a different format for negative numbers, then the number should be made positive
		// before getting formatted. The format string itself will contain formatting that denotes a negative number and
		// this formatting will end up in the prefix or suffix. Commonly if there is a negative format specified, the
		// number will get surrounded by parenthesis instead of showing it with a minus sign.
		if fullFormat.negativeFormatExpectsPositive {
			floatVal = math.Abs(floatVal)
		}
		numberFormat = fullFormat.negativeFormat
	} else {
		numberFormat = fullFormat.zeroFormat
	}

	// When showPercent is true, multiply the number by 100.
	// The percent sign will be in the prefix or suffix already, so it does not need to be added in this function.
	// The number format itself will be the same as any other number format once the value is multiplied by 100.
	if numberFormat.showPercent {
		floatVal = 100 * floatVal
	}

	// Only the most common format strings are supported here.
	// Eventually this switch needs to be replaced with a more general solution.
	// Some of these "supported" formats should have thousand separators, but don't get them since Go fmt
	// doesn't have a way to request thousands separators.
	// The only things that should be supported here are in the array formattingCharacters,
	// everything else has been stripped out before and will be placed in the prefix or suffix.
	// The formatting characters can have non-formatting characters mixed in with them and those should be maintained.
	// However, at this time we fail to parse those formatting codes and they get replaced with "General"
	var formattedNum string
	switch numberFormat.reducedFormatString {
	case "general":
		// prefix, showPercent, and suffix cannot apply to the general format
		// The logic for showing numbers when the format is "general" is much more complicated than the rest of these.
		generalFormatted, err := formatter.generalNumericScientific(cellValue)
		if err != nil {
			return rawValue, nil
		}
		formatter.setSeparators(&generalFormatted, formatter.decimalSeparator, "")
		return generalFormatted, nil
	case "@": // String is "@"
		formattedNum = cellValue
	case "0":
		formattedNum = fmt.Sprintf("%.0f", floatVal)
	case "#,##0":
		formattedNum = fmt.Sprintf("%.0f", floatVal)
		formatter.setSeparators(&formattedNum, formatter.decimalSeparator, "")
	case "###0.0", "0.0":
		formattedNum = fmt.Sprintf("%.1f", floatVal)
		formatter.setSeparators(&formattedNum, formatter.decimalSeparator, "")
	case "###0.00", "0.00":
		formattedNum = fmt.Sprintf("%.2f", floatVal)
		formatter.setSeparators(&formattedNum, formatter.decimalSeparator, "")
	case "###0.000", "0.000":
		formattedNum = fmt.Sprintf("%.3f", floatVal)
		formatter.setSeparators(&formattedNum, formatter.decimalSeparator, "")
	case "###0.0000", "0.0000":
		formattedNum = fmt.Sprintf("%.4f", floatVal)
		formatter.setSeparators(&formattedNum, formatter.decimalSeparator, "")
	case "#,##0.0":
		formattedNum = fmt.Sprintf("%.1f", floatVal)
		formatter.setSeparators(&formattedNum, formatter.decimalSeparator, formatter.thousandSeparator)
	case "#,##0.00": // Float is "0.00"
		formattedNum = fmt.Sprintf("%.2f", floatVal)
		formatter.setSeparators(&formattedNum, formatter.decimalSeparator, formatter.thousandSeparator)
	case "#,##0.000":
		formattedNum = fmt.Sprintf("%.3f", floatVal)
		formatter.setSeparators(&formattedNum, formatter.decimalSeparator, formatter.thousandSeparator)
	case "#,##0.0000":
		formattedNum = fmt.Sprintf("%.4f", floatVal)
		formatter.setSeparators(&formattedNum, formatter.decimalSeparator, formatter.thousandSeparator)
	case "0.00e+00", "##0.0e+0":
		digits := strings.IndexRune(numberFormat.reducedFormatString, '+')
		if -1 == digits {
			digits = 1
		} else {
			digits = 2
		}
		if formatter.allowScientific {
			formattedNum := fmt.Sprintf("%."+strconv.Itoa(digits)+"e", floatVal)
			formatter.setSeparators(&formattedNum, formatter.decimalSeparator, "")
			return formattedNum, nil
		} else {
			formatter.setSeparators(&rawValue, formatter.decimalSeparator, "")
			return rawValue, nil
		}
	case "":
		// Do nothing.
	default:
		return rawValue, nil
	}
	return numberFormat.prefix + formattedNum + numberFormat.suffix, nil
}

func (formatter *excelFormatter) setSeparators(renderedNumber *string, decimalSeparator string, thousandSeparator string) {
	var signLen int
	var fracPosition int
	fracPosition = strings.IndexRune(*renderedNumber, '.')
	if -1 == fracPosition {
		fracPosition = len(*renderedNumber)
	}
	if len(*renderedNumber) > 0 && (*renderedNumber)[0] == '-' {
		signLen = 1
	}
	for i := fracPosition - 3; i > signLen; i -= 3 {
		*renderedNumber = (*renderedNumber)[0:i] + "#" + (*renderedNumber)[i:]
	}
	if len(decimalSeparator) > 0 && decimalSeparator != "." {
		*renderedNumber = strings.Replace(*renderedNumber, ".", decimalSeparator, -1)
	}
	*renderedNumber = strings.Replace(*renderedNumber, "#", thousandSeparator, -1)
	return
}

func (formatter *excelFormatter) generalNumericScientific(value string) (string, error) {
	if strings.TrimSpace(value) == "" {
		return "", nil
	}
	f, err := strconv.ParseFloat(value, 64)
	if err != nil {
		return value, err
	}
	if formatter.allowScientific {
		absF := math.Abs(f)
		// When using General format, numbers that are less than 1e-9 (0.000000001) and greater than or equal to
		// 1e11 (100,000,000,000) should be shown in scientific notation.
		// Numbers less than the number after zero, are assumed to be zero.
		if (absF >= math.SmallestNonzeroFloat64 && absF < minNonScientificNumber) || absF >= maxNonScientificNumber {
			return strconv.FormatFloat(f, 'E', -1, 64), nil
		}
	}
	// This format (fmt="f", prec=-1) will prevent padding with zeros and will never switch to scientific notation.
	// However, it will show more than 11 characters for very precise numbers, and this cannot be changed.
	// You could also use fmt="g", prec=11, which doesn't pad with zeros and allows the correct precision,
	// but it will use scientific notation on numbers less than 1e-4. That value is hardcoded in Go and cannot be
	// configured or disabled.
	return strconv.FormatFloat(f, 'f', -1, 64), nil
}

func (formatter *excelFormatter) parseTime(value string, fullFormat *parsedNumberFormat) (string, error) {
	f, err := strconv.ParseFloat(value, 64)
	if err != nil {
		return value, err
	}
	val := TimeFromExcelTime(f, formatter.date1904)
	format := fullFormat.numFmt
	if formatter.dateFixedFormat != "" {
		format = formatter.dateFixedFormat
	}
	// Replace Excel placeholders with Go time placeholders.
	// For example, replace yyyy with 2006. These are in a specific order,
	// due to the fact that m is used in month, minute, and am/pm. It would
	// be easier to fix that with regular expressions, but if it's possible
	// to keep this simple it would be easier to maintain.
	// Full-length month and days (e.g. March, Tuesday) have letters in them that would be replaced
	// by other characters below (such as the 'h' in March, or the 'd' in Tuesday) below.
	// First we convert them to arbitrary characters unused in Excel Date formats, and then at the end,
	// turn them to what they should actually be.
	// Based off: http://www.ozgrid.com/Excel/CustomFormats.htm
	replacements := []struct {
		xltime string
		gotime string
		limit  int
	}{
		{"yyyy", "2006", 1},
		{"yy", "06", 1},
		{"mmmmm", "%%%%%", 1},
		{"mmmm", "%%%%", 1},
		{"mmm", "%%%", 1},
		{"dddd", "&&&&", 1},
		{"ddd", "&&&", 1},
		{"dd", "02", 1},
		{"d", "2", 1},
		{"mmm", "Jan", 1},
		{"mmss", "0405", 1},
		{"ss", "05", 1},
		{"mm:", "04:", 1},
		{":mm", ":04", 1},
		{"mm", "01", 1},
		{"am/pm", "pm", 1},
		{"m/", "1/", 1},
		//{"%%%%", "January", 1},
		//{"&&&&", "Monday", 1},
		//{"&&&", "Mon", 1},
		{"\\ ", " ", -1},
		{"\\,", ",", -1},
	}
	// It is the presence of the "am/pm" indicator that determins
	// if this is a 12 hour or 24 hours time format, not the
	// number of 'h' characters.
	if is12HourTime(format) {
		format = strings.Replace(format, "hh", "03", 1)
		format = strings.Replace(format, "h", "3", 1)
	} else {
		format = strings.Replace(format, "hh", "15", 1)
		format = strings.Replace(format, "h", "15", 1)
	}
	for _, repl := range replacements {
		format = strings.Replace(format, repl.xltime, repl.gotime, repl.limit)
	}
	// If the hour is optional, strip it out, along with the
	// possible dangling colon that would remain.
	if val.Hour() < 1 {
		format = strings.Replace(format, "]:", "]", 1)
		format = strings.Replace(format, "[03]", "", 1)
		format = strings.Replace(format, "[3]", "", 1)
		format = strings.Replace(format, "[15]", "", 1)
	} else {
		format = strings.Replace(format, "[3]", "3", 1)
		format = strings.Replace(format, "[15]", "15", 1)
	}
	weekdayId := int(val.Weekday())
	monthId := int(val.Month())
	for _, repl := range []struct {
		macro string
		value string
	}{
		{"%%%%%", formatter.i18n.monthNamesPasv[monthId]},
		{"%%%%", formatter.i18n.monthNames[monthId]},
		{"%%%", formatter.i18n.monthNames3[monthId]},
		{"&&&&", formatter.i18n.weekdayNames[weekdayId]},
		{"&&&", formatter.i18n.weekdayNames3[weekdayId]},
	} {
		format = strings.Replace(format, repl.macro, repl.value, -1)
	}
	return val.Format(format), nil
}

// isTimeFormat checks whether an Excel format string represents a time.Time.
// This function is now correct, but it can detect time format strings that cannot be correctly handled by parseTime()
func isTimeFormat(format string) bool {
	var foundTimeFormatCharacters bool
	for i := 0; i < len(format); i++ {
		curReducedFormat := format[i:]
		switch curReducedFormat[0] {
		case '\\', '_':
			// If there is a slash, skip the next character, and add it to the prefix
			// If there is an underscore, skip the next character, but don't add it to the prefix
			if len(curReducedFormat) > 1 {
				i++
			}
		case '*':
			// Asterisks are used to repeat the next character to fill the full cell width.
			// There isn't really a cell size in this context, so this will be ignored.
		case '"':
			// If there is a quote skip to the next quote, and add the quoted characters to the prefix
			endQuoteIndex := strings.Index(curReducedFormat[1:], "\"")
			if endQuoteIndex == -1 {
				// This is not any type of valid format.
				return false
			}
			i += endQuoteIndex + 1
		case '$', '-', '+', '/', '(', ')', ':', '!', '^', '&', '\'', '~', '{', '}', '<', '>', '=', ' ', '.':
			// These symbols are allowed to be used as literal without escaping
		case ',':
			// This is not documented in the XLSX spec as far as I can tell, but Excel and Numbers will include
			// commas in number formats without escaping them, so this should be supported.
		default:
			foundInThisLoop := false
			for _, special := range timeFormatCharacters {
				if strings.HasPrefix(curReducedFormat, special) {
					foundTimeFormatCharacters = true
					foundInThisLoop = true
					i += len(special) - 1
					break
				}
			}
			if foundInThisLoop {
				continue
			}
			if curReducedFormat[0] == '[' {
				// For number formats, this code would happen above in a case '[': section.
				// However, for time formats it must happen after looking for occurrences in timeFormatCharacters
				// because there are a few time formats that can be wrapped in brackets.

				// Brackets can be currency annotations (e.g. [$$-409])
				// color formats (e.g. [color1] through [color56], as well as [red] etc.)
				// conditionals (e.g. [>100], the valid conditionals are =, >, <, >=, <=, <>)
				bracketIndex := strings.Index(curReducedFormat, "]")
				if bracketIndex == -1 {
					// This is not any type of valid format.
					return false
				}
				i += bracketIndex
				continue
			}
			// Symbols that don't have meaning, aren't in the exempt literal characters, and aren't escaped are invalid.
			// The string could still be a valid number format string.
			return false
		}
	}
	// If the string doesn't have any time formatting characters, it could technically be a time format, but it
	// would be a pretty weak time format. A valid time format with no time formatting symbols will also be a number
	// format with no number formatting symbols, which is essentially a constant string that does not depend on the
	// cell's value in anyway. The downstream logic will do the right thing in that case if this returns false.
	return foundTimeFormatCharacters
}

// is12HourTime checks whether an Excel time format string is a 12
// hours form.
func is12HourTime(format string) bool {
	return strings.Contains(format, "am/pm") || strings.Contains(format, "AM/PM") || strings.Contains(format, "a/p") || strings.Contains(format, "A/P")
}

func shiftJulianToNoon(julianDays, julianFraction float64) (float64, float64) {
	switch {
	case -0.5 < julianFraction && julianFraction < 0.5:
		julianFraction += 0.5
	case julianFraction >= 0.5:
		julianDays += 1
		julianFraction -= 0.5
	case julianFraction <= -0.5:
		julianDays -= 1
		julianFraction += 1.5
	}
	return julianDays, julianFraction
}

// Return the integer values for hour, minutes, seconds and
// nanoseconds that comprised a given fraction of a day.
// values would round to 1 us.
func fractionOfADay(fraction float64) (hours, minutes, seconds, nanoseconds int) {

	const (
		c1us  = 1e3
		c1s   = 1e9
		c1day = 24 * 60 * 60 * c1s
	)

	frac := int64(c1day*fraction + c1us/2)
	nanoseconds = int((frac%c1s)/c1us) * c1us
	frac /= c1s
	seconds = int(frac % 60)
	frac /= 60
	minutes = int(frac % 60)
	hours = int(frac / 60)
	return
}

func julianDateToGregorianTime(part1, part2 float64) time.Time {
	part1I, part1F := math.Modf(part1)
	part2I, part2F := math.Modf(part2)
	julianDays := part1I + part2I
	julianFraction := part1F + part2F
	julianDays, julianFraction = shiftJulianToNoon(julianDays, julianFraction)
	day, month, year := doTheFliegelAndVanFlandernAlgorithm(int(julianDays))
	hours, minutes, seconds, nanoseconds := fractionOfADay(julianFraction)
	return time.Date(year, time.Month(month), day, hours, minutes, seconds, nanoseconds, time.UTC)
}

// By this point generations of programmers have repeated the
// algorithm sent to the editor of "Communications of the ACM" in 1968
// (published in CACM, volume 11, number 10, October 1968, p.657).
// None of those programmers seems to have found it necessary to
// explain the constants or variable names set out by Henry F. Fliegel
// and Thomas C. Van Flandern.  Maybe one day I'll buy that jounal and
// expand an explanation here - that day is not today.
func doTheFliegelAndVanFlandernAlgorithm(jd int) (day, month, year int) {
	l := jd + 68569
	n := (4 * l) / 146097
	l = l - (146097*n+3)/4
	i := (4000 * (l + 1)) / 1461001
	l = l - (1461*i)/4 + 31
	j := (80 * l) / 2447
	d := l - (2447*j)/80
	l = j / 11
	m := j + 2 - (12 * l)
	y := 100*(n-49) + i + l
	return d, m, y
}

// Convert an excelTime representation (stored as a floating point number) to a time.Time.
func TimeFromExcelTime(excelTime float64, date1904 bool) time.Time {
	var date time.Time
	var wholeDaysPart = int(excelTime)
	// Excel uses Julian dates prior to March 1st 1900, and
	// Gregorian thereafter.
	if wholeDaysPart <= 61 {
		const OFFSET1900 = 15018.0
		const OFFSET1904 = 16480.0
		var date time.Time
		if date1904 {
			date = julianDateToGregorianTime(MJD_0, excelTime+OFFSET1904)
		} else {
			date = julianDateToGregorianTime(MJD_0, excelTime+OFFSET1900)
		}
		return date
	}
	var floatPart = excelTime - float64(wholeDaysPart)
	if date1904 {
		date = excel1904Epoc
	} else {
		date = excel1900Epoc
	}
	durationPart := time.Duration(nanosInADay * floatPart)
	return date.AddDate(0, 0, wholeDaysPart).Add(durationPart)
}
