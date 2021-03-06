package tablescanner

type tI18n struct {
	decimalSeparator  string // not yet used
	thousandSeparator string // not yet used
	weekdayNames      [7]string
	weekdayNames3     [7]string
	monthNames        [13]string
	monthNamesPasv    [13]string
	monthNames3       [13]string
	numFmtDefaults    map[int]string
	numFmtSystem      map[string]string
}

var numFmtI18n = map[string]*tI18n{
	"en": {
		decimalSeparator:  ".",
		thousandSeparator: ",",
		weekdayNames: [7]string{
			"Sunday",
			"Monday",
			"Tuesday",
			"Wednesday",
			"Thursday",
			"Friday",
			"Saturday",
		},
		weekdayNames3: [7]string{
			"Sun",
			"Mon",
			"Tue",
			"Wed",
			"Thu",
			"Fri",
			"Sat",
		},
		monthNames: [13]string{
			"",
			"January",
			"February",
			"March",
			"April",
			"May",
			"June",
			"July",
			"August",
			"September",
			"October",
			"November",
			"December",
		},
		monthNamesPasv: [13]string{
			"",
			"January",
			"February",
			"March",
			"April",
			"May",
			"June",
			"July",
			"August",
			"September",
			"October",
			"November",
			"December",
		},
		monthNames3: [13]string{
			"",
			"Jan",
			"Feb",
			"Mar",
			"Apr",
			"May",
			"Jun",
			"Jul",
			"Aug",
			"Sep",
			"Oct",
			"Nov",
			"Dec",
		},
		numFmtSystem: map[string]string{
			"[$-F800]": "dddd, mmmm dd, yyyy",
			"[$-FC19]": "dddd, mmmm dd, yyyy",
		},
		numFmtDefaults: map[int]string{
			0:  "general",
			1:  "0",
			2:  "0.00",
			3:  "#,##0",
			4:  "#,##0.00",
			5:  "#,##0",
			6:  "#,##0",
			7:  "#,##0.00",
			8:  "#,##0.00",
			9:  "0%",        // need to multiply by 100
			10: "0.00%",     // need to multiply by 100
			11: "0.00e+00",  // exp with 1-9.(9) significand
			12: "#\" \"?/?", // quoted inline string containing space, 1-dig frac
			13: "# ??/??",   // quoted inline string containing space, 2-dig frac
			14: "m/d/yyyy",
			15: "d-mmm-yy",
			16: "d-mmm",
			17: "mmm-yy",
			18: "h:mm am/pm", // <am/pm> causes <h> to be 12-h type
			19: "h:mm:ss am/pm",
			20: "h:mm",
			21: "h:mm:ss",
			22: "m/d/yyyy h:mm",
			23: "general",
			24: "general",
			25: "general",
			26: "general",
			27: "m/d/yyyy",
			28: "m/d/yyyy",
			29: "m/d/yyyy",
			30: "m/d/yyyy",
			31: "m/d/yyyy",
			32: "h:mm:ss",
			33: "h:mm:ss",
			34: "h:mm:ss",
			35: "h:mm:ss",
			36: "m/d/yyyy",
			37: "#,##0",
			38: "#,##0",
			39: "#,##0.00",
			40: "#,##0.00",
			41: "#,##0",
			42: "#,##0",
			43: "#,##0.00",
			44: "#,##0.00",
			45: "mm:ss",
			46: "[h]:mm:ss", /// =frac(value)+24h*int(value)
			47: "mm:ss.0",   // secs with 1/10
			48: "##0.0e+0",  // exp with 10-99.(9) significand
			49: "@",
			50: "m/d/yyyy",
			51: "m/d/yyyy",
			52: "m/d/yyyy",
			53: "m/d/yyyy",
			54: "m/d/yyyy",
			55: "m/d/yyyy",
			56: "m/d/yyyy",
			57: "m/d/yyyy",
			58: "m/d/yyyy",
			59: "0",
			60: "0.00",
			61: "#,##0",
			62: "#,##0.00",
			63: "#,##0",
			64: "#,##0",
			65: "#,##0.00",
			66: "#,##0.00",
			67: "0%",
			68: "0.00%",
			69: "#\" \"?/?",
			70: "# ??/??",
			71: "m/d/yyyy",
			72: "m/d/yyyy",
			73: "d-mmm-yy",
			74: "d-mmm",
			75: "mmm-yy",
			76: "h:mm",
			77: "h:mm:ss",
			78: "m/d/yyyy h:mm",
			79: "h:mm",
			80: "[h]:mm:ss",
			81: "mm:ss.0",
		},
	},
	"ru": {
		decimalSeparator:  ",",
		thousandSeparator: "\xC2\xA0",
		monthNames: [13]string{
			"",
			"Январь",
			"Февраль",
			"Март",
			"Апрель",
			"Май",
			"Июнь",
			"Июль",
			"Август",
			"Сентябрь",
			"Октябрь",
			"Ноябрь",
			"Декабрь",
		},
		monthNamesPasv: [13]string{
			"",
			"января",
			"февраля",
			"марта",
			"апреля",
			"мая",
			"июня",
			"июля",
			"августа",
			"сентября",
			"октября",
			"ноября",
			"декабря",
		},
		monthNames3: [13]string{
			"",
			"янв",
			"фев",
			"мар",
			"апр",
			"май",
			"июн",
			"июл",
			"авг",
			"сен",
			"окт",
			"ноя",
			"дек",
		},
		weekdayNames: [7]string{
			"Воскресенье",
			"Понедельник",
			"Вторник",
			"Среда",
			"Четверг",
			"Пятница",
			"Суббота",
		},
		weekdayNames3: [7]string{
			"ВС",
			"ПН",
			"ВТ",
			"СР",
			"ЧТ",
			"ПТ",
			"СБ",
		},
		numFmtSystem: map[string]string{
			"[$-F800]": "d mmmmm yyyy г.",
			"[$-FC19]": "d mmmmm yyyy г.",
		},
		numFmtDefaults: map[int]string{
			0:  "general",
			1:  "0",
			2:  "0.00",
			3:  "#,##0",
			4:  "#,##0.00",
			5:  "#,##0",
			6:  "#,##0",
			7:  "#,##0.00",
			8:  "#,##0.00",
			9:  "0%",
			10: "0.00%",
			11: "0.00e+00",
			12: "#\" \"?/?",
			13: "# ??/??",
			14: "dd.mm.yyyy",
			15: "dd.mmm.yy",
			16: "dd.mmm",
			17: "mmm.yy",
			18: "h:mm am/pm",
			19: "h:mm:ss am/pm",
			20: "h:mm",
			21: "h:mm:ss",
			22: "dd.mm.yyyy h:mm",
			23: "general",
			24: "general",
			25: "general",
			26: "general",
			27: "dd.mm.yyyy",
			28: "dd.mm.yyyy",
			29: "dd.mm.yyyy",
			30: "dd.mm.yyyy",
			31: "dd.mm.yyyy",
			32: "h:mm:ss",
			33: "h:mm:ss",
			34: "h:mm:ss",
			35: "h:mm:ss",
			36: "dd.mm.yyyy",
			37: "#,##0",
			38: "#,##0",
			39: "#,##0.00",
			40: "#,##0.00",
			41: "#,##0",
			42: "#,##0",
			43: "#,##0.00",
			44: "#,##0.00",
			45: "mm:ss",
			46: "[h]:mm:ss",
			47: "mm:ss.0",
			48: "##0.0e+0",
			49: "@",
			50: "dd.mm.yyyy",
			51: "dd.mm.yyyy",
			52: "dd.mm.yyyy",
			53: "dd.mm.yyyy",
			54: "dd.mm.yyyy",
			55: "dd.mm.yyyy",
			56: "dd.mm.yyyy",
			57: "dd.mm.yyyy",
			58: "dd.mm.yyyy",
			59: "0",
			60: "0.00",
			61: "#,##0",
			62: "#,##0.00",
			63: "#,##0",
			64: "#,##0",
			65: "#,##0.00",
			66: "#,##0.00",
			67: "0%",
			68: "0.00%",
			69: "#\" \"?/?",
			70: "# ??/??",
			71: "dd.mm.yyyy",
			72: "dd.mm.yyyy",
			73: "d.mmm.yy",
			74: "d.mmm",
			75: "mmm.yy",
			76: "h:mm",
			77: "h:mm:ss",
			78: "dd.mm.yyyy h:mm",
			79: "h:mm",
			80: "[h]:mm:ss",
			81: "mm:ss.0",
		},
	},
}
