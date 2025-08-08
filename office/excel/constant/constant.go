package constant

import "time"

const (
	LetterMax           = 26             // excel字母坐标最大值
	DefaultHeaderHeight = 24             // 默认表头高度
	DefaultSheetName    = "Sheet1"       // 默认工作表名称
	DefaultLine         = 1              // 默认表头行号
	DefaultExcelType    = "xlsx"         // 默认Excel文件类型
	MaxRowNum           = 100000         // 导出默认最大数据量
	MaxErrNum           = 100            // 错误最大量
	ResultExpTime       = 24 * time.Hour // 过期时间(24小时)
)
