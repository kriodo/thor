package datetime

import (
	"fmt"
	"github.com/kriodo/thor/tool"
	"github.com/xuri/excelize/v2"
	"strconv"
	"strings"
	"time"
)

// TimestampToTime 时间戳(秒)转time.Time
func TimestampToTime(val int64) time.Time {
	return time.Unix(val, 0)
}

// TimestampToDate 时间戳(秒)转日期
func TimestampToDate(val int64, df DateFormat) string {
	return time.Unix(val, 0).Format(df.String())
}

// TimestampFormat 时间戳(秒)按照标准格式化  eg：年月日时间戳 -> 年月时间戳
func TimestampFormat(val int64, df DateFormat) int64 {
	if val <= 0 {
		return 0
	}
	return DateToTimestamp(TimestampToDate(val, df), df)
}

// TimestampToZhMonth 时间戳(秒)转日期 (中文月份)
func TimestampToZhMonth(ts int64) string {
	timeStr := time.Unix(ts, 0).Format("1")
	monthMap := map[string]string{
		"1":  "一",
		"2":  "二",
		"3":  "三",
		"4":  "四",
		"5":  "五",
		"6":  "六",
		"7":  "七",
		"8":  "八",
		"9":  "九",
		"10": "十",
		"11": "十一",
		"12": "十二",
	}
	return monthMap[timeStr]
}

// ExcelDateToTime excel日期转datetime 45352 => time.Time(注意: 时区为UTC)
func ExcelDateToTime(excelDate string) (time.Time, error) {
	f, err := strconv.ParseFloat(excelDate, 64)
	if err != nil {
		return time.Time{}, err
	}
	if f <= 0 {
		return time.Time{}, nil
	}
	return excelize.ExcelDateToTime(f, false)
}

// AnyToTimestamp 各种日期格式转时间戳（秒）
func AnyToTimestamp(date string) int64 {
	if tool.ContainsEnglishWord(date) {
		tt, err := time.Parse("Jan-06", date)
		if err != nil {
			return 0
		}
		return tt.Unix()

	}
	var (
		dateLen   = tool.LenChar(date)
		tt        = time.Time{}
		err       error
		formatArr []string
	)

	switch dateLen {
	case 4:
		formatArr = []string{"2006"}
	case 5: // excel时间
		tt, err = ExcelDateToTime(date)
		if err != nil && tt.Unix() > 0 {
			return tt.Unix()
		}
		formatArr = []string{"20061", "2006年"}
	case 6:
		formatArr = []string{"200601", "2006-1", "2006/1", "2006.1", "200601", "1-2-06", "1/2/06", "1.2.06"}
	case 7:
		formatArr = []string{"2006-01", "2006/01", "2006.01", "2006-01", "2006/01", "2006.01", "2006年1月"}
	case 8:
		formatArr = []string{"20060102", "2006/1/2", "2006-1-2", "2006.1.2", "01-02-06", "15:04:05", "20060102", "2006年01月"}
	case 9:
		formatArr = []string{"2006年1月2号", "2006年1月2日"}
	case 10:
		formatArr = []string{"2006-01-02", "2006/01/02", "2006.01.02", "2006-01-02",
			"2006/01/02", "2006.01.02", "2006年01月2号", "2006年1月02号",
			"2006年01月2日", "2006年1月02日"}
	case 11:
		formatArr = []string{"2006年01月02号", "2006年01月02日", "20060102 15", "2006年01月02日"}
	case 12:
		formatArr = []string{"1/2/06 00:00", "1-2-06 00:00", "1.2.06 00:00", "06/1/2 00:00", "06-1-2 00:00", "06-1-2 00:00"}
	case 13:
		formatArr = []string{"2006-01-02 15", "2006/01/02 15", "2006.01.02 15"}
	case 14:
		formatArr = []string{"20060102 15:04"}
	case 15:
		formatArr = []string{"2006年01月02日 15时"}
	case 16:
		formatArr = []string{"2006-01-02 15:04", "2006/01/02 15:04", "2006.01.02 15:04"}
	case 17:
		formatArr = []string{"20060102 15:04:05"}
	case 18:
		formatArr = []string{"2006年01月02日 15时04分"}
	case 19:
		formatArr = []string{"2006-01-02 15:04:05", "2006/01/02 15:04:05", "2006.01.02 15:04:05"}
	case 21:
		formatArr = []string{"2006年01月02日 15时04分05秒"}
	default:
		return 0
	}
	for _, v := range formatArr {
		tt, err = DateToTime(date, DateFormat(v))
		if err != nil && tt.Unix() <= 0 {
			continue
		}
		return tt.Unix()
	}
	return 0

}

// FormatStringToRangeDate 解析字符串区间日期（年月）如：2006.01-2006.06
func FormatStringToRangeDate(dateStr string) (startTime int64, endTime int64, err error) {
	if dateStr == "" {
		return 0, 0, fmt.Errorf("日期必填")
	}
	timeArr := strings.Split(dateStr, "-")
	dataLen := len(timeArr)
	switch dataLen {
	case 1:
		startTime = AnyToTimestamp(dateStr)
		if startTime <= 0 {
			return 0, 0, fmt.Errorf("日期格式错误")
		}
	case 2: // 2024-06|2024.06-2024.06|2024/06-2024/06
		dataLen1 := tool.LenChar(timeArr[0])
		if dataLen1 < 4 { // 第一个值必须>4位
			return 0, 0, fmt.Errorf("日期格式错误")
		} else if dataLen1 == 4 {
			startTime = AnyToTimestamp(dateStr)
			if startTime <= 0 {
				return 0, 0, fmt.Errorf("日期格式错误")
			}
		} else if dataLen1 <= 7 {
			startTime = AnyToTimestamp(timeArr[0])
			if startTime <= 0 {
				return 0, 0, fmt.Errorf("日期格式错误")
			}
			endTime = AnyToTimestamp(timeArr[1])
			if endTime <= 0 {
				return 0, 0, fmt.Errorf("日期格式错误")
			}
		} else {
			return 0, 0, fmt.Errorf("日期格式错误")
		}
	case 3: // 2024.06-2024-06|2024-6-2024.6
		dataLen1 := tool.LenChar(timeArr[0])
		if dataLen1 < 4 {
			return 0, 0, fmt.Errorf("日期格式错误")
		}
		if dataLen1 == 4 { // 2024-6-2024.6
			startTime = AnyToTimestamp(timeArr[0] + "-" + timeArr[1])
			if startTime <= 0 {
				return 0, 0, fmt.Errorf("日期格式错误")
			}
			endTime = AnyToTimestamp(timeArr[2])
			if endTime <= 0 {
				return 0, 0, fmt.Errorf("日期格式错误")
			}
		} else { // 2024.06-2024-06
			startTime = AnyToTimestamp(timeArr[0])
			if startTime <= 0 {
				return 0, 0, fmt.Errorf("日期格式错误")
			}
			endTime = AnyToTimestamp(timeArr[1] + "-" + timeArr[2])
			if endTime <= 0 {
				return 0, 0, fmt.Errorf("日期格式错误")
			}
		}
	case 4: // 2024-06-2024-06|2024-6-2024-6
		dataLen1 := tool.LenChar(timeArr[0])
		dataLen2 := tool.LenChar(timeArr[2])
		if dataLen1 != 4 && dataLen2 != 4 {
			return 0, 0, fmt.Errorf("日期格式错误")
		}
		startTime = AnyToTimestamp(timeArr[0] + "-" + timeArr[1])
		if startTime <= 0 {
			return 0, 0, fmt.Errorf("日期格式错误")
		}
		endTime = AnyToTimestamp(timeArr[2] + "-" + timeArr[3])
		if endTime <= 0 {
			return 0, 0, fmt.Errorf("日期格式错误")
		}
	default:
		return 0, 0, fmt.Errorf("日期格式错误")
	}
	if startTime <= 0 && endTime <= 0 {
		return 0, 0, fmt.Errorf("日期格式错误")
	}
	return
}
