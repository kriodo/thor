package datetime

import "time"

// DateToTime 日期转time.Time
func DateToTime(date string, df DateFormat) (time.Time, error) {
	loc, err := time.LoadLocation(time.Local.String())
	if err != nil {
		return time.Time{}, err
	}
	return time.ParseInLocation(df.String(), date, loc)
}

// DateToTimestamp 日期转时间戳(秒)
func DateToTimestamp(date string, df DateFormat) int64 {
	if date == "" {
		return 0
	}
	ts, err := DateToTime(date, df)
	if err != nil {
		return 0
	}

	return ts.Unix()
}

// Date2MilliTimestamp 日期转时间戳(毫秒)
func Date2MilliTimestamp(date string, df DateFormat) int64 {
	if date == "" {
		return 0
	}
	ts, err := DateToTime(date, df)
	if err != nil {
		return 0
	}
	return ts.UnixMilli()
}
