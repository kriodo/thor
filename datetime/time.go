package datetime

import "time"

// StartOfDate 获取日期的开始日期
func StartOfDate(val time.Time) time.Time {
	return time.Date(val.Year(), val.Month(), val.Day(), 0, 0, 0, 0, time.Local)
}

// EndOfDate 获取日期的结束日期
func EndOfDate(val time.Time) time.Time {
	return time.Date(val.Year(), val.Month(), val.Day(), 23, 59, 59, 0, time.Local)
}

// StartOfToday 今天的开始日期
func StartOfToday() time.Time {
	return StartOfDate(time.Now())
}

// EndOfToday 今天的结束日期
func EndOfToday() time.Time {
	return EndOfDate(time.Now())
}

// StartOfWeek 星期的开始日期
func StartOfWeek(val time.Time) time.Time {
	var (
		add int
		w   = val.Weekday()
	)
	switch w {
	case time.Sunday:
		add = -6
	default:
		add = 1 - int(w)
	}
	return StartOfDate(AddDate(val, 0, 0, add))
}

// EndOfWeek 星期的结束日期
func EndOfWeek(val time.Time) time.Time {
	return StartOfWeek(AddDate(val, 0, 0, 7)).Add(time.Second * -1)
}

// StartOfThisWeek start of this week
func StartOfThisWeek() time.Time {
	return StartOfWeek(time.Now())
}

// EndOfThisWeek end of this week
func EndOfThisWeek() time.Time {
	return EndOfWeek(time.Now())
}

// StartOfMonth start of month
func StartOfMonth(val time.Time) time.Time {
	return time.Date(val.Year(), val.Month(), 1, 0, 0, 0, 0, time.Local)
}

// EndOfMonth end of month
func EndOfMonth(val time.Time) time.Time {
	return EndOfDate(StartOfMonth(AddDate(val, 0, 1, 0)).Add(time.Second * -1))
}

// StartOfThisMonth start of this month
func StartOfThisMonth() time.Time {
	return StartOfMonth(time.Now())
}

// EndOfThisMonth end of this month
func EndOfThisMonth() time.Time {
	return EndOfMonth(time.Now())
}

// StartOfYear start of year
func StartOfYear(val time.Time) time.Time {
	return time.Date(val.Year(), 1, 1, 0, 0, 0, 0, time.Local)
}

// EndOfYear end of year
func EndOfYear(val time.Time) time.Time {
	return StartOfYear(AddDate(val, 1, 0, 0)).Add(time.Second * -1)
}

// StartOfThisYear start of this year
func StartOfThisYear() time.Time {
	return StartOfYear(time.Now())
}

// EndOfThisYear end of this year
func EndOfThisYear() time.Time {
	return EndOfYear(time.Now())
}

// AddDate 日期加值(原AddDate临t,界时间会有歧义)
func AddDate(tt time.Time, year, month, day int) time.Time {
	// 先跳到目标月的1号
	targetDate := tt.AddDate(year, month, -tt.Day()+1)
	// 获取目标月的临界值
	targetDay := targetDate.AddDate(0, 1, -1).Day()
	// 对比临界值与源日期值，取最小的值
	if targetDay > tt.Day() {
		targetDay = tt.Day()
	}
	// 最后用目标月的1号加上目标值和入参的天数
	targetDate = targetDate.AddDate(0, 0, targetDay-1+day)
	return targetDate
}
