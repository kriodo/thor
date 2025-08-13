package header

// Style 表头样式
type Style struct {
	FontFamily          string  // 字体类型
	FontSize            float64 // 字体大小
	FontColor           string  // 字体颜色
	FontBold            bool    // 字体加粗
	BorderColor         string  // 边框颜色
	AlignmentHorizontal string  // 水平居中
	FillColor           string  // 填充颜色
	CustomNumFmt        *string
}

// GetExportDefaultStyle 获取默认导出的excel表头样式
func GetExportDefaultStyle() *Style {
	return &Style{
		FontFamily:  "微软雅黑",
		FontSize:    11,
		BorderColor: "000000",
		FillColor:   "DBE2F6",
	}
}
