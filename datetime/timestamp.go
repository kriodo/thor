package datetime

import (
	"fmt"
	"github.com/kriodo/thor/tool"
	"github.com/xuri/excelize/v2"
	"math"
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

// NDaysAfter  n天后的end of day 比如当前时间: 2023-03-10 15:00:00
func NDaysAfter(n int) time.Time {
	return EndOfToday().AddDate(0, 0, n)
}

// MonthTimestampRange 获取两个时间戳(月)之间的月份时间戳(秒) 如：【2006-01,2007-02]
func MonthTimestampRange(startVal, endVal int64) []int64 {
	// 格式化时间戳
	var (
		start = TimestampFormat(startVal, YYYYMM_0)
		end   = TimestampFormat(endVal, YYYYMM_0)
	)
	if start > end {
		return []int64{}
	}
	if start == end {
		return []int64{start}
	}
	var (
		startTime = time.Unix(start, 0)
		dates     []int64
		counter   int
	)
	for {
		ts := startTime.Unix()
		dates = append(dates, ts)
		if ts == end {
			break
		}
		if counter >= 1000 { // 显示数据量别整太大
			break
		}
		startTime = AddDate(startTime, 0, 1, 0)
		counter++
	}

	return dates
}

// MonthDateRange 获取两个时间戳(月)之间的月份时间戳(秒) 如：【2006-01,2007-02]
func MonthDateRange(startVal, endVal int64) []string {
	vals := MonthTimestampRange(startVal, endVal)
	list := make([]string, 0, len(vals))
	for _, val := range vals {
		list = append(list, TimestampToDate(val, YYYYMM_0))
	}
	return list
}

// DayTimestampRange 获取两个时间戳(日)之间的日期时间戳(秒) 如：【2006-01-01,2006-01-02]
func DayTimestampRange(startVal, endVal int64) []int64 {
	// 格式化时间戳
	var (
		start = TimestampFormat(startVal, YYYYMMDD_0)
		end   = TimestampFormat(endVal, YYYYMMDD_0)
	)
	if start > end {
		return []int64{}
	}
	if start == end {
		return []int64{start}
	}
	var (
		startTime = time.Unix(start, 0)
		days      []int64
		counter   int
	)

	for {
		ts := startTime.Unix()
		days = append(days, ts)
		if ts == end {
			break
		}
		if counter >= 1000 { // 显示数据量别整太大
			break
		}
		startTime = AddDate(startTime, 0, 0, 1)
		counter++
	}

	return days
}

// DayDateRange 获取两个时间戳(日)之间的日期时间戳(秒) 如：【2006-01-01,2006-01-02]
func DayDateRange(startVal, endVal int64) []string {
	vals := DayTimestampRange(startVal, endVal)
	list := make([]string, 0, len(vals))
	for _, val := range vals {
		list = append(list, TimestampToDate(val, YYYYMMDD_0))
	}
	return list
}

// GetBetweenDates 根据开始日期和结束日期计算出时间段内所有（年月日）日期如：2020-01-01
// Example:
//
//	d1 := GetBetweenDates("2023-03-01", "2023-03-03")  //output: [2023-03-01 2023-03-02 2023-03-03]
func GetBetweenDates(startDate, endDate string) []string {
	if startDate == endDate {
		return []string{startDate}
	}
	d := []string{}
	timeFormatTpl := string(YYYYMMDD_HHMMSS_0)
	if len(timeFormatTpl) != len(startDate) {
		timeFormatTpl = timeFormatTpl[0:len(startDate)]
	}
	date, err := time.Parse(timeFormatTpl, startDate)
	if err != nil {
		// 时间解析，异常
		return d
	}
	date2, err := time.Parse(timeFormatTpl, endDate)
	if err != nil {
		// 时间解析，异常
		return d
	}
	if date2.Before(date) {
		// 如果结束时间小于开始时间，异常
		return d
	}
	// 输出日期格式固定
	timeFormatTpl = string(YYYYMMDD_0)
	date2Str := date2.Format(timeFormatTpl)
	d = append(d, date.Format(timeFormatTpl))
	for {
		date = date.AddDate(0, 0, 1)
		dateStr := date.Format(timeFormatTpl)
		d = append(d, dateStr)
		if dateStr == date2Str {
			break
		}
	}
	return d
}

// GetBetweenMonths 根据开始日期和结束日期计算出时间段内所有（年月）日期
func GetBetweenMonths(startDate, endDate string) []string {
	if startDate == endDate {
		return []string{startDate}
	}
	tArr := []string{}
	t1, err := DateToTime(startDate, YYYYMM_0)
	if err != nil {
		return tArr
	}
	t2, err := DateToTime(endDate, YYYYMM_0)
	if err != nil {

		return tArr
	}
	tArr = append(tArr, t1.Format(string(YYYYMM_0)))
	var i int = 1
	for {
		t := t1.AddDate(0, i, 0)
		tArr = append(tArr, t.Format(string(YYYYMM_0)))
		if t == t2 {
			break
		}
		i++
		if i > 200 { // 防止死循环
			break
		}
	}

	return tArr
}

// GetBetweenMonthsFormat 根据开始日期和结束日期, 以指定格式计算出时间段内所有（年月）日期
func GetBetweenMonthsFormat(startDate string, endDate string, format DateFormat) []string {
	if startDate == endDate {
		return []string{startDate}
	}
	tArr := []string{}
	t1, err := DateToTime(startDate, format)
	if err != nil {
		return tArr
	}
	t2, err := DateToTime(endDate, format)
	if err != nil {
		return tArr
	}
	tArr = append(tArr, t1.Format(string(format)))
	var i int = 1
	for {
		t := t1.AddDate(0, i, 0)
		tArr = append(tArr, t.Format(string(format)))
		if t == t2 {
			break
		}
		i++
		if i > 200 { // 防止死循环
			break
		}
	}

	return tArr
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

// 解析 "Mar-24" 这样的日期格式，并将其解释为 "2024年3月"
func parseMonthYear(dateStr string) (time.Time, error) {
	// 定义月份名称到月份数字的映射
	monthNames := map[string]int{
		"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4,
		"May": 5, "Jun": 6, "Jul": 7, "Aug": 8,
		"Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
	}

	// 分割日期字符串
	parts := strings.Split(dateStr, "-")
	if len(parts) != 2 {
		return time.Time{}, fmt.Errorf("invalid date format: %s", dateStr)
	}

	// 解析月份
	month, ok := monthNames[parts[0]]
	if !ok {
		return time.Time{}, fmt.Errorf("invalid month: %s", parts[0])
	}

	// 获取当前年份
	currentYear := time.Now().Year()
	// 将年份转换为字符串
	yearStr := strconv.Itoa(currentYear)

	// 提取年份的前两位
	firstTwoDigits := yearStr[:2]

	// 解析年份
	year, err := strconv.Atoi(firstTwoDigits + parts[1])
	if err != nil {
		return time.Time{}, fmt.Errorf("invalid year: %s", parts[1])
	}

	// 创建时间对象
	t := time.Date(year, time.Month(month), 1, 0, 0, 0, 0, time.Local)

	return t, nil
}

// LastMonthDate 获取上个月的日期 如：2006-02 12:01:02 ->  2006-01 12:01:02
func LastMonthDate(t int64, df DateFormat) string {
	return time.Unix(t, 0).AddDate(0, -1, 0).Format(string(df))
}

// NextMonthDate 获取下个月的日期 如：2006-02 12:01:02 ->  2006-03 12:01:02
func NextMonthDate(t int64, df DateFormat) string {
	return time.Unix(t, 0).AddDate(0, 1, 0).Format(string(df))
}

// AddDateOfStartMonth 日期加值-月初
func AddDateOfStartMonth(t time.Time, year, month, day int) time.Time {
	tt := AddDate(t, year, month, day)
	return time.Date(tt.Year(), tt.Month(), 1, 0, 0, 0, 0, tt.Location())
}

// FillTime 填充当前时间 如果只有月份 填充当前年
func FillTime(input string) (string, error) {
	if _, err := strconv.Atoi(input); err == nil {
		if len(input) == 1 {
			input = fmt.Sprintf("0%s", input)
		}
		currentYear := time.Now().Year()
		input = fmt.Sprintf("%v-%s", currentYear, input)
	}

	input = strings.TrimSpace(input)
	year := strconv.Itoa(time.Now().Year())

	// 如果输入只有一个数字，补充年份
	if len(input) == 1 {
		input = year + "-" + input
	}

	// 检查是否包含月份信息
	if strings.Contains(input, "月") {
		// 检查是否包含年份
		if !strings.Contains(input, "年") {
			input = year + "年" + input
		}
		return input, nil
	}

	// 如果输入长度为2，补充年份
	if len(input) == 2 {
		input = year + "-" + input
	}

	// 如果输入长度为6，补充年份和连接符
	if len(input) == 6 {
		input = year + "-" + input[:2] + "-" + input[2:]
	}

	// 解析时间，检查是否合法
	_, err := time.Parse("2006-01", input)
	if err != nil {
		return "", fmt.Errorf("日期格式错误：%s", input)
	}

	return input, nil
}

// 判断是否是周末
func IsWeekend(ts int64) bool {
	weekday := time.Unix(ts, 0).Weekday()
	return weekday == time.Sunday || weekday == time.Saturday
}

// 计算两个日期相差几天
func DaysBetween2Date(format DateFormat, date1Str, date2Str string) (int, error) {
	// 将字符串转化为Time格式
	date1, err := time.ParseInLocation(format.String(), date1Str, time.Local)
	if err != nil {
		return 0, err
	}
	// 将字符串转化为Time格式
	date2, err := time.ParseInLocation(format.String(), date2Str, time.Local)
	if err != nil {
		return 0, err
	}
	//计算相差天数
	differ := math.Ceil(date1.Sub(date2).Hours() / 24)
	return int(differ), nil
}

// DaysBetween2Unix 计算两个时间戳相差几天
func DaysBetween2Unix(ts1, ts2 int64) int {
	// 将字符串转化为Time格式
	date1 := time.Unix(ts1, 0)
	date2 := time.Unix(ts2, 0)
	//计算相差天数
	differ := math.Ceil(date1.Sub(date2).Hours() / 24)
	return int(differ)
}

// 将毫秒数转换为人类可读的时间描述，例如“1小时30分钟”。
func FormatMsDesc(ms int64) string {
	duration := time.Duration(ms) * time.Millisecond
	hours := duration / time.Hour
	duration -= hours * time.Hour

	minutes := duration / time.Minute
	duration -= minutes * time.Minute

	seconds := duration / time.Second
	duration -= seconds * time.Second
	ms = duration.Milliseconds()
	result := ""
	if hours > 0 {
		result += fmt.Sprintf("%d小时", hours)
	}
	if minutes > 0 {
		result += fmt.Sprintf("%d分钟", minutes)
	}
	if seconds > 0 {
		result += fmt.Sprintf("%d秒", seconds)
	}
	if ms > 0 || result == "" {
		result += fmt.Sprintf("%d毫秒", ms)
	}
	return result
}

// 获取指定时间戳的月份天数 到日
func TimestampMonthDays(ts int64) int64 {
	tsTime := time.Unix(ts, 0)
	start := StartOfMonth(tsTime)
	end := EndOfMonth(tsTime)
	allDays := GetBetweenDates(TimestampToDate(start.Unix(), YYYYMMDD_0), TimestampToDate(end.Unix(), YYYYMMDD_0))
	return int64(len(allDays))
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
