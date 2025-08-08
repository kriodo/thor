package excel

//var (
//	Letters = []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
//)
//
//// IndexLetter 根据表头获取excel横坐标标识
//func IndexLetter(h []*Header) {
//	var (
//		index int // 下标
//		count int // LetterMax倍数计数器
//	)
//	// 获取letter
//	getLetter := func(s string, num int) string {
//		for j := 0; j < num; j++ {
//			s += Letters[index]
//		}
//		return s
//	}
//	for k := range h {
//		h[k].Abscissa = getLetter(Letters[index], count)
//		if index >= header.LetterMax-1 {
//			index = 0 // 重置下标key
//			count++
//			continue
//		}
//		index++
//	}
//}
//
//type ExportSetting struct {
//	NowSheetName string    // 工作表名称
//	FileName  string    // 文件名称
//	Line      int       // 纵坐标行号
//	Height    int       // 行高
//	Header    []*Header // 表头
//}
//
//// 自定义表头结构
//type CustomHeader struct {
//	Header  []*Header  // 表头
//	RowData [][]string // 数据
//	File    *excelize.File
//	eopts   exportOption
//}
//
//// 表头行参数
//type Header struct {
//	Name     string // 名称
//	Abscissa string // 横坐标字母
//	Width    int    // 行宽
//}
//
//func NewExport(h []*Header, rd [][]string, opts ...ExportOption) (*CustomHeader, error) {
//	ch := &CustomHeader{
//		Header:  h,
//		RowData: rd,
//		File:    excelize.NewFile(),
//		eopts:   defaultExportOptions(),
//	}
//	// 循环调用opts
//	for _, opt := range opts {
//		opt.exportApply(&ch.eopts)
//	}
//
//	return ch, nil
//}
//
//// 需要设置默认值参数
//type exportOption struct {
//	Line      int
//	Height    float64
//	FileName  string
//	NowSheetName string
//}
//
//type ExportOption interface {
//	exportApply(option *exportOption)
//}
//
//type funcOption struct {
//	ef func(option *exportOption)
//}
//
//func (fdo *funcOption) exportApply(do *exportOption) {
//	fdo.ef(do)
//}
//
//func newFuncOption(ef func(option *exportOption)) *funcOption {
//	return &funcOption{
//		ef: ef,
//	}
//}
//
//// 默认参数
//func defaultExportOptions() exportOption {
//	return exportOption{
//		Line:      header.DefaultLine,
//		Height:    header.DefaultHeaderHeight,
//		FileName:  tool.GetUUID() + "." + header.DefaultExcelType,
//		NowSheetName: header.DefaultSheetName,
//	}
//}
//
//// 设置表头行号
//func WithHeaderLine(line int) ExportOption {
//	return newFuncOption(func(option *exportOption) {
//		option.Line = line
//	})
//}
//
//// 设置表头行高
//func WithHeaderHeight(heigth float64) ExportOption {
//	return newFuncOption(func(option *exportOption) {
//		option.Height = heigth
//	})
//}
//
//// 设置工作表名称
//func WithHeaderSheetName(sheetName string) ExportOption {
//	return newFuncOption(func(option *exportOption) {
//		option.NowSheetName = sheetName
//	})
//}
//
//// 设置文件名称
//func WithFileName(fileName string) ExportOption {
//	return newFuncOption(func(option *exportOption) {
//		option.FileName = fileName
//	})
//}
//
//func (ch *CustomHeader) SetHeader() error {
//	// 创建一个工作表
//	index, err := ch.File.NewSheet(ch.eopts.NowSheetName)
//	if err != nil {
//		return errors.New("创建工作表失败")
//	}
//
//	// 设置工作簿的默认工作表
//	ch.File.SetActiveSheet(index)
//
//	// 给表头增加坐标
//	IndexLetter(ch.Header)
//
//	// 设置表头单元格
//	length := len(ch.Header)
//	for _, header := range ch.Header {
//		err := ch.File.SetCellValue(ch.eopts.NowSheetName, fmt.Sprintf("%s%d", header.Abscissa, ch.eopts.Line), header.Name)
//		if err != nil {
//			return err
//		}
//	}
//
//	// 设置行高
//	err = ch.File.SetRowHeight(ch.eopts.NowSheetName, ch.eopts.Line, ch.eopts.Height)
//	if err != nil {
//		return errors.New("设置表头行高失败")
//	}
//
//	// 设置单元格边框
//	border := []excelize.Border{
//		{Type: "top", Style: 2, Color: "#cccccc"},
//		{Type: "left", Style: 2, Color: "#cccccc"},
//		{Type: "right", Style: 2, Color: "#cccccc"},
//		{Type: "bottom", Style: 2, Color: "#cccccc"},
//	}
//	// 定义标题行单元格样式
//	headerStyle, err := ch.File.NewStyle(&excelize.Style{
//		Border: border,
//		Fill: excelize.Fill{
//			Type:    "pattern",
//			Color:   []string{"#e6e6e6"},
//			Pattern: 1, // 图案填充
//		},
//		Font: &excelize.Font{Bold: true},
//		Alignment: &excelize.Alignment{
//			Horizontal:      "left", // 水平效果
//			Indent:          0,
//			JustifyLastLine: false,
//			ReadingOrder:    0,
//			RelativeIndent:  0,
//			ShrinkToFit:     false,
//			TextRotation:    0,
//			Vertical:        "center", // 垂直效果
//			WrapText:        false,
//		},
//		Protection:   nil,
//		NumFmt:       0,
//		CustomNumFmt: nil,
//		NegRed:       false,
//	},
//	)
//	if err != nil {
//		return errors.New("设置单元格样式错误")
//	}
//
//	// 为标题行设置样式
//	err = ch.File.SetCellStyle(ch.eopts.NowSheetName, fmt.Sprintf("%s%d", "A", ch.eopts.Line), fmt.Sprintf("%s%d", ch.Header[length-1].Abscissa, ch.eopts.Line), headerStyle)
//	if err != nil {
//		return errors.New("设置单元格样式错误")
//	}
//
//	return nil
//}
