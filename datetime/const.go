package datetime

// DateFormat 标准时间格式
type DateFormat string

const (
	YYYY              DateFormat = "2006"
	HHMMSS            DateFormat = "15:04:05"
	YYYYMM_0          DateFormat = "2006-01"
	YYYYMMDD_0        DateFormat = "2006-01-02"
	YYYYMMDD_HH_0     DateFormat = "2006-01-02 15"
	YYYYMMDD_HHMM_0   DateFormat = "2006-01-02 15:04"
	YYYYMMDD_HHMMSS_0 DateFormat = "2006-01-02 15:04:05"

	YYYYMM_1          DateFormat = "2006/01"
	YYYYMMDD_1        DateFormat = "2006/01/02"
	YYYYMMDD_HH_1     DateFormat = "2006/01/02 15"
	YYYYMMDD_HHMM_1   DateFormat = "2006/01/02 15:04"
	YYYYMMDD_HHMMSS_1 DateFormat = "2006/01/02 15:04:05"

	YYYYMM_3          DateFormat = "2006.01"
	YYYYMMDD_3        DateFormat = "2006.01.02"
	YYYYMMDD_HH_3     DateFormat = "2006.01.02 15"
	YYYYMMDD_HHMM_3   DateFormat = "2006.01.02 15:04"
	YYYYMMDD_HHMMSS_3 DateFormat = "2006.01.02 15:04:05"

	YYYYMM_4          DateFormat = "200601"
	YYYYMMDD_4        DateFormat = "20060102"
	YYYYMMDD_HH_4     DateFormat = "20060102 15"
	YYYYMMDD_HHMM_4   DateFormat = "20060102 15:04"
	YYYYMMDD_HHMMSS_4 DateFormat = "20060102 15:04:05"

	YYYY_5            DateFormat = "2006年"
	YYYYMM_5          DateFormat = "2006年01月"
	YYYYMMDD_5        DateFormat = "2006年01月02日"
	YYYYMMDD_HH_5     DateFormat = "2006年01月02日 15时"
	YYYYMMDD_HHMM_5   DateFormat = "2006年01月02日 15时04分"
	YYYYMMDD_HHMMSS_5 DateFormat = "2006年01月02日 15时04分05秒"

	YYYYMMDD_EXCEL DateFormat = "01-02-06"
)

func (df DateFormat) String() string {
	return string(df)
}
