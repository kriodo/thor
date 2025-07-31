package datetime

import (
	"fmt"
	"math"
	"time"
)

// IsWeekend 判断是否是周末
func IsWeekend(ts int64) bool {
	weekday := time.Unix(ts, 0).Weekday()
	return weekday == time.Sunday || weekday == time.Saturday
}

// FormatMsDesc 将毫秒数转换为人类可读的时间描述，例如“1小时30分钟”。
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

// DiffDays 计算两个日期相差天数
func DiffDays(start, end time.Time) int {
	//计算相差天数
	differ := math.Ceil(start.Sub(end).Hours() / 24)
	return int(differ)
}

// DiffDaysByDate 计算两个日期相差天数
func DiffDaysByDate(start, end string, df DateFormat) (int, error) {
	// 将字符串转化为Time格式
	val1, err := DateToTime(start, df)
	if err != nil {
		return 0, err
	}
	// 将字符串转化为Time格式
	val2, err := DateToTime(end, df)
	if err != nil {
		return 0, err
	}
	//计算相差天数
	differ := DiffDays(val1, val2)
	return int(differ), nil
}

// DiffDaysByTimestamp  计算两个时间戳相差几天
func DiffDaysByTimestamp(start, end int64) int {
	// 将字符串转化为Time格式
	val1 := time.Unix(start, 0)
	val2 := time.Unix(end, 0)
	//计算相差天数
	differ := DiffDays(val1, val2)
	return int(differ)
}

// BetweenDaysForTimestamp 获取两个时间戳(秒)之间的日期时间戳(秒) 如：【2006-01-01,2006-01-02]
func BetweenDaysForTimestamp(startVal, endVal int64) []int64 {
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

// BetweenDaysForDate 获取两个时间戳(日)之间的日期时间戳(秒) 如：【2006-01-01,2006-01-02]
func BetweenDaysForDate(startVal, endVal int64) []string {
	vals := BetweenDaysForTimestamp(startVal, endVal)
	list := make([]string, 0, len(vals))
	for _, val := range vals {
		list = append(list, TimestampToDate(val, YYYYMMDD_0))
	}
	return list
}

// BetweenMonthsForDate 根据开始日期和结束日期计算出时间段内所有（年月）日期
func BetweenMonthsForDate(startDate, endDate string) []string {
	if startDate == endDate {
		return []string{startDate}
	}
	var tArr []string
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

// BetweenMonthByTimestamp 获取两个时间戳(月)之间的月份时间戳(秒) 如：【2006-01,2007-02]
func BetweenMonthByTimestamp(startVal, endVal int64) []int64 {
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
