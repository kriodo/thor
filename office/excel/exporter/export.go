package exporter

import (
	"fmt"
	"github.com/kriodo/thor/office/excel/constant"
	"github.com/kriodo/thor/office/excel/header"
	"github.com/kriodo/thor/office/tool"
	"github.com/xuri/excelize/v2"
)

// Exporter 导出
type Exporter struct {
	NowSheetName string                 // 当前sheet名称
	file         *excelize.File         // 文件
	headerMap    map[string]*headerInfo // 表头map

	headerTree         []*header.Header    // 表头tree
	styleId            int                 // 样式id
	maxLevel           int                 // 层级
	fieldKeyMap        map[string]int      // 字段对应的索引值
	fieldKeyList       []string            // 有效字段排序列表
	err                error               // 错误
	startColumn        int                 // 起始表头列: 默认=1
	maxColumn          int                 // 最大列数
	maxColumnLetter    string              // 最大列数对应的字母坐标
	startLine          uint                // 起始表头行: 默认=1
	rowStyleList       map[string]int      // 设置指定单元格样式
	lineRowStyleList   map[int]int         // 设置行内容样式（一整行） 行号:StyleId
	validationDropList map[string][]string // 设置指定单元格序列数据验证
}

// 表头数据
type headerInfo struct {
	headerTree         []*header.Header    // 表头tree
	styleId            int                 // 样式id
	maxLevel           int                 // 层级
	fieldKeyMap        map[string]int      // 字段对应的索引值
	fieldKeyList       []string            // 有效字段排序列表
	err                error               // 错误
	startColumn        int                 // 起始表头列: 默认=1
	maxColumn          int                 // 最大列数
	maxColumnLetter    string              // 最大列数对应的字母坐标
	startLine          uint                // 起始表头行: 默认=1
	rowStyleList       map[string]int      // 设置指定单元格样式
	lineRowStyleList   map[int]int         // 设置行内容样式（一整行） 行号:StyleId
	validationDropList map[string][]string // 设置指定单元格序列数据验证
}

// NewExporter new一个导出处理器
func NewExporter(sheetName string) (*Exporter, error) {
	err := tool.CheckSheetName(sheetName)
	if err != nil {
		return nil, err
	}
	file := excelize.NewFile()
	// 创建一个工作表
	index, err := file.NewSheet(sheetName)
	if err != nil {
		return nil, fmt.Errorf("导出创建sheet失败: %+v", err)
	}
	if sheetName != constant.DefaultSheetName {
		err = file.DeleteSheet(constant.DefaultSheetName)
		if err != nil {
			return nil, fmt.Errorf("导出删除默认sheet失败: %+v", err)
		}
	}
	// 设置默认sheet
	file.SetActiveSheet(index)
	er := &Exporter{
		NowSheetName: sheetName,
		file:         file,
		headerMap:    make(map[string]*headerInfo, 10),
	}
	er.headerMap[er.NowSheetName] = &headerInfo{}
	return er, nil
}

// NewExporterWithFile new一个导出处理器: 自带文件
func NewExporterWithFile(file *excelize.File) (*Exporter, error) {
	sheetLen := len(file.GetSheetList())
	if sheetLen == 0 {
		return nil, fmt.Errorf("表格无sheet数据")
	}
	er := &Exporter{
		file:         file,
		NowSheetName: file.GetSheetList()[0],
		fieldKeyMap:  make(map[string]int),
		headerMap:    make(map[string]*headerInfo, 10),
	}
	er.headerMap[er.NowSheetName] = &headerInfo{}
	return er, nil
}

// AddSheet 增加一个sheet
func (er *Exporter) AddSheet(sheetName string) (*Exporter, error) {
	err := tool.CheckSheetName(sheetName)
	if err != nil {
		return nil, err
	}
	if tool.InStringArray(sheetName, er.file.GetSheetList()) {
		return nil, fmt.Errorf(sheetName + "已存在，无法创建")
	}
	_, err = er.file.NewSheet(sheetName)
	if err != nil {
		return er, fmt.Errorf("创建sheet失败: %s %+v", sheetName, err)
	}
	er.NowSheetName = sheetName
	return er, nil
}

// ChangeSheet 切换sheet
func (er *Exporter) ChangeSheet(sheetName string) *Exporter {
	er.NowSheetName = sheetName
	return er
}

func (er *Exporter) clear() {
	er.fieldKeyMap = make(map[string]int)
	er.fieldKeyList = []string{}
	er.rowStyleList = make(map[string]int)
	er.lineRowStyleList = make(map[int]int)
	er.validationDropList = make(map[string][]string)
	er.maxColumn = 0
	er.maxColumnLetter = ""
}

// SetTree 设置tree数据 [此模式id、pid无需填写]
func (er *Exporter) SetTree(tree []*header.Header) *Exporter {
	er.clear()
	header.FormatTree(tree, 1, nil)
	er.maxLevel = header.MaxLevel(tree, 1)
	er.headerTree = tree
	if er.styleId <= 0 {
		er.defaultStyle()
	}
	if er.startLine < 1 {
		er.startLine = 1
	}
	if er.startLine > 1 {
		er.maxLevel = er.maxLevel + er.startLine - 1
	}
	_, er.err = er.formatHeader(er.headerTree, er.startColumn, er.startLine)
	// 设置最后表头数据宽度
	er.setHeaderWidth(er.headerTree, er.startColumn)
	return er
}

func (er *Exporter) setHeaderWidth(tree []*header.Header, xIndex int) int {
	for _, header := range tree {
		if header.IsLastLevel {
			x := tool.IndexToLetter(xIndex)
			w := float64(tool.LenChar(header.Title))*1.7 + 8
			er.file.SetColWidth(er.NowSheetName, x, x, w)
			xIndex = xIndex + 1
		} else {
			xIndex = er.setHeaderWidth(header.Children, xIndex)
		}
	}
	return xIndex
}

// SetListV2 设置list数据 [此模式pkey必须填写]
func (er *Exporter) SetListV2(headers []*header.Header) *Exporter {
	var (
		newHeaders []*header.Header
		headerL1   []*header.Header
		headerMap  = make(map[string]int64)
		id         int64
	)
	for _, header := range headers {
		id++
		header.Id = id
		headerL1 = append(headerL1, header)
		headerMap[header.FieldKey] = id
	}
	for _, header := range headers {
		header.Pid = headerMap[header.Pkey]
		newHeaders = append(newHeaders, header)
	}

	return er.SetList(newHeaders)
}

// SetList 设置list数据 [此模式id、pid必须填写]
func (er *Exporter) SetList(headers []*header.Header) *Exporter {
	er.clear()
	err := header.Validation(headers)
	if err != nil {
		er.err = err
		return er
	}
	if er.err != nil {
		return er
	}
	tree := header.List2Tree(headers, 0)
	header.FormatTree(tree, 1, nil)
	er.maxLevel = header.MaxLevel(tree, 1)
	er.headerTree = tree
	if er.styleId <= 0 {
		er.defaultStyle()
	}
	if er.startLine < 1 {
		er.startLine = 1
	}
	if er.startLine > 1 {
		er.maxLevel = er.maxLevel + er.startLine - 1
	}
	_, er.err = er.formatHeader(er.headerTree, er.startColumn, er.startLine)
	// 设置最后表头数据宽度
	er.setHeaderWidth(er.headerTree, er.startColumn)
	return er
}

// SetStartIndex 设置起始列（此函数必须在处理表头相关函数之后，且设置数据之前之间，否则会出现数据问题）
func (er *Exporter) SetStartIndex(index int) *Exporter {
	er.startColumn = index
	return er
}

// SetDataStartLine 设置数据起始行（此函数必须在处理表头相关函数之后，且设置数据之前之间，否则会出现数据问题）
func (er *Exporter) SetDataStartLine(index uint) *Exporter {
	er.startLine = index
	return er
}

// SetListData 设置数据 [此模式数据需要按照顺序排序]
func (er *Exporter) SetListData(rows [][]interface{}) *Exporter {
	for index, row := range rows {
		rowAddr, err := excelize.JoinCellName("A", index+er.maxLevel+1)
		if err != nil {
			er.err = fmt.Errorf("set data join cell name failed. %+v", err)
			return er
		}
		if err = er.file.SetSheetRow(er.NowSheetName, rowAddr, &row); err != nil {
			er.err = fmt.Errorf("set data set sheet row failed. %+v", err)
			return er
		}
	}
	return er
}

// SetRowStyle 设置列样式
func (er *Exporter) SetRowStyle(rowStyleList map[string]int) *Exporter {
	er.rowStyleList = rowStyleList
	return er
}

// SetValidationString 设置序号验证
func (er *Exporter) SetValidationString(validationDropList map[string][]string) *Exporter {
	er.validationDropList = validationDropList
	return er
}

// SetLineRowStyle 设置表头样式
func (er *Exporter) SetLineRowStyle(lineRowStyle map[int]int) *Exporter {
	er.lineRowStyleList = lineRowStyle
	return er
}

// SetMapData 设置数据 [此模式只需要对应的key附上值即可]
func (er *Exporter) SetMapData(dataMap []map[string]interface{}) *Exporter {
	rows := make([][]interface{}, 0, len(dataMap))
	for _, m := range dataMap {
		row := make([]interface{}, len(er.fieldKeyList)+er.startColumn)
		for _, key := range er.fieldKeyList {
			data, exi1 := m[key]
			index, exi2 := er.fieldKeyMap[key]
			if exi1 && exi2 {
				row[index] = data
			}
		}
		rows = append(rows, row)
	}
	indexToTitle := er.reversalFieldKeyMap(er.fieldKeyMap)
	for index, row := range rows {
		// 按照行设置样式
		lineRowStyleId, _ := er.lineRowStyleList[index]
		for key, val := range row {
			letter := tool.IndexToLetter(key)
			rowAddr, err := excelize.JoinCellName(letter, index+er.maxLevel+1)
			if err = er.file.SetCellValue(er.NowSheetName, rowAddr, val); err != nil {
				er.err = fmt.Errorf("set data set sheet row failed. error=%+v row=%+v", err, index)
				return er
			}

			// 验证列内容样式
			if getRowStyleId, ok := er.rowStyleList[indexToTitle[key]]; ok {
				if err = er.file.SetCellStyle(er.NowSheetName, rowAddr, rowAddr, getRowStyleId); err != nil {
					er.err = fmt.Errorf("set data row sytle cell failed. error=%+v row=%+v", err, index)
					return er
				}
			}

			// 设置数据验证
			if getValidationBody, ok := er.validationDropList[indexToTitle[key]]; ok {
				dv := excelize.NewDataValidation(true)
				dv.SetSqref(fmt.Sprintf("%s:%s", rowAddr, rowAddr))
				err = dv.SetDropList(getValidationBody)
				if err != nil {
					er.err = fmt.Errorf("set data SetDropList row failed. error=%+v row=%+v", err, index)
					return er
				}
				er.file.AddDataValidation(er.NowSheetName, dv)
			}

			if lineRowStyleId > 0 {
				if err = er.file.SetCellStyle(er.NowSheetName, rowAddr, rowAddr, lineRowStyleId); err != nil {
					er.err = fmt.Errorf("set data line row sytle cell failed. error=%+v row=%+v", err, index)
					return er
				}
			}
		}
	}
	return er
}

// SetMapDataNoStyle 设置数据不设置样式（速度快点） [此模式只需要对应的key附上值即可]
func (er *Exporter) SetMapDataNoStyle(dataMap []map[string]interface{}) *Exporter {
	rows := make([][]interface{}, 0, len(dataMap))
	for _, m := range dataMap {
		row := make([]interface{}, len(er.fieldKeyList)+er.startColumn)
		for _, key := range er.fieldKeyList {
			data, exi1 := m[key]
			index, exi2 := er.fieldKeyMap[key]
			if exi1 && exi2 {
				row[index] = data
			}
		}
		rows = append(rows, row)
	}
	er.SetListData(rows)
	return er
}

type RowData struct {
	Value     interface{} // 值
	StyleId   int         // 样式id
	ValueType             // 值类型
}

type ValueType int

const (
	ValueCellDefault ValueType = 0 //默认
	ValueCellString  ValueType = 1 //字符串
)

// SetMapDataWithStyle 设置数据 + 设置样式 [此模式只需要对应的key附上值即可]
func (er *Exporter) SetMapDataWithStyle(dataMap []map[string]*RowData) *Exporter {
	rows := make([][]*RowData, 0, len(dataMap))
	for _, m := range dataMap {
		row := make([]*RowData, len(er.fieldKeyList)+er.startColumn)
		for _, key := range er.fieldKeyList {
			data, exi1 := m[key]
			index, exi2 := er.fieldKeyMap[key]
			if exi1 && exi2 {
				row[index] = data
			}
		}
		rows = append(rows, row)
	}
	for index, row := range rows {
		for key, info := range row {
			if info == nil {
				continue
			}
			letter := tool.IndexToLetter(key)
			rowAddr, err := excelize.JoinCellName(letter, index+er.maxLevel+1)
			switch info.ValueType {
			case ValueCellString:
				err = er.file.SetCellStr(er.NowSheetName, rowAddr, tool.Any2String(info.Value))
			default:
				err = er.file.SetCellValue(er.NowSheetName, rowAddr, info.Value)
			}
			if err != nil {
				er.err = fmt.Errorf("set data set sheet row value failed. error=%+v row=%+v", err, index)
				return er
			}
			if info.StyleId > 0 {
				if err = er.file.SetCellStyle(er.NowSheetName, rowAddr, rowAddr, info.StyleId); err != nil {
					er.err = fmt.Errorf("set data row sytle cell failed. error=%+v row=%+v", err, index)
					return er
				}
			}
		}
	}
	return er
}

// reversalFieldKeyMap 反转表头
func (er *Exporter) reversalFieldKeyMap(fieldKeyMap map[string]int) map[int]string {
	var fieldList = make(map[int]string)
	for key, val := range fieldKeyMap {
		fieldList[val] = key
	}
	return fieldList
}

func (er *Exporter) SetStyle(headerStyle *header.Style) *Exporter {
	if headerStyle == nil {
		return er
	}
	// 定义边框样式
	border := []excelize.Border{
		{Type: "top", Style: 1, Color: headerStyle.BorderColor},
		{Type: "left", Style: 1, Color: headerStyle.BorderColor},
		{Type: "right", Style: 1, Color: headerStyle.BorderColor},
		{Type: "bottom", Style: 1, Color: headerStyle.BorderColor},
	}
	// 定义标题行单元格样式
	styleId, err := er.file.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold:   false,
			Family: headerStyle.FontFamily,
			Size:   headerStyle.FontSize,
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{headerStyle.FillColor},
			Pattern: 1,
		},
		Border:       border,
		CustomNumFmt: headerStyle.CustomNumFmt,
		Alignment: &excelize.Alignment{
			Vertical:   "center",
			Horizontal: "center",
		}},
	)
	if err != nil {
		er.err = fmt.Errorf("set cell style failed. %+v", err)
		return er
	}
	er.styleId = styleId
	return er
}

func (er *Exporter) GetHeaderStyleId(headerStyle *header.Style) int {
	if headerStyle == nil {
		return 0
	}
	// 定义边框样式
	border := []excelize.Border{
		{Type: "top", Style: 1, Color: headerStyle.BorderColor},
		{Type: "left", Style: 1, Color: headerStyle.BorderColor},
		{Type: "right", Style: 1, Color: headerStyle.BorderColor},
		{Type: "bottom", Style: 1, Color: headerStyle.BorderColor},
	}
	// 定义标题行单元格样式
	alignmentHorizontal := "center"
	if headerStyle.AlignmentHorizontal != "" {
		alignmentHorizontal = headerStyle.AlignmentHorizontal
	}
	styleId, err := er.file.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold:   headerStyle.FontBold,
			Family: headerStyle.FontFamily,
			Size:   headerStyle.FontSize,
			Color:  headerStyle.FontColor,
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{headerStyle.FillColor},
			Pattern: 1,
		},
		Border:       border,
		CustomNumFmt: headerStyle.CustomNumFmt,
		Alignment: &excelize.Alignment{
			Vertical:   "center",
			Horizontal: alignmentHorizontal,
		}},
	)
	if err != nil {
		er.err = fmt.Errorf("set cell style failed. %+v", err)
		return 0
	}
	return styleId
}

// File 获取文件file
func (er *Exporter) File() *excelize.File {
	return er.file
}

// GetFieldKeyMap 获取表头字段对应的索引值
func (er *Exporter) GetFieldKeyMap() map[string]int {
	return er.fieldKeyMap
}

func (er *Exporter) GetFieldEnCol(filedKey string) string {
	i, ok := er.fieldKeyMap[filedKey]
	if !ok {
		return ""
	}
	return tool.IndexToLetter(i)
}

// GetFieldKeyList 获取表头有效字段排序列表
func (er *Exporter) GetFieldKeyList() []string {
	return er.fieldKeyList
}

// 获取错误
func (er *Exporter) Error() error {
	return er.err
}

// SaveAs 文件存储
func (er *Exporter) SaveAs(filePath string) error {
	if er.err != nil {
		return er.err
	}
	if err := er.file.SaveAs(filePath); err != nil {
		return err
	}
	return nil
}

// SetStringStyle 设置列的文本格式
// FieldKeys 字段key
// StartLine 开始行号: <=0时候默认为1
// EndLine 结束行号: <=0时候默认为100
func (er *Exporter) SetStringStyle(fieldKeys []string, startLine, endLine int) error {
	if len(fieldKeys) == 0 {
		return nil
	}
	if startLine <= 0 {
		startLine = 1
	}
	if endLine <= 0 {
		startLine = 100
	}
	fieldKeys = tool.UniqueString(fieldKeys)
	// 创建一个新的样式
	strStyleId, err := er.File().NewStyle(&excelize.Style{NumFmt: 49}) // 49="@" 表示文本格式
	if err != nil {
		return err
	}
	for _, key := range fieldKeys {
		if index, exi := er.fieldKeyMap[key]; exi {
			keyAbc := tool.IndexToLetter(index)
			if keyAbc == "" {
				continue
			}
			err = er.File().SetCellStyle(er.NowSheetName, fmt.Sprintf("%s%d", keyAbc, startLine), fmt.Sprintf("%s%d", keyAbc, endLine), strStyleId)
			if err != nil {
				return err
			}
		}
	}
	return nil
}

// 设置多级表头，会根据层级自动合并
func (er *Exporter) formatHeader(headers []*header.Header, index, line int) (int, error) {
	var err error
	for _, v := range headers {
		err = er.left2right(v, index, line)
		if err != nil {
			return 0, err
		}
		if len(v.Children) > 0 {
			if er.startLine < 1 {
				er.startLine = 1
			}
			_, err = er.formatHeader(v.Children, index, line+1)
			if err != nil {
				return 0, err
			}
		} else {
			if er.maxLevel > v.Level {
				err = er.up2down(v, index, line)
				if err != nil {
					return 0, err
				}
			}
		}
		if v.IsLastLevel {
			er.fieldKeyMap[v.FieldKey] = index
			er.fieldKeyList = append(er.fieldKeyList, v.FieldKey)
		}
		index += header.ChildrenLen(v.Children)
	}
	return index, nil
}

// 左右合并
func (er *Exporter) left2right(val *header.Header, index int, line int) error {
	// 左右合并单元格
	leftLetter := tool.ExcelIndexLetter(index)
	left, err := excelize.JoinCellName(leftLetter, line)
	if err != nil {
		return fmt.Errorf("left right val join cell name 1 failed. val: %+v err: %+v", val, err)
	}
	if err = er.file.SetCellValue(er.NowSheetName, left, val.Title); err != nil {
		return fmt.Errorf("left right val set cell value failed. val: %+v err: %+v", val, err)
	}
	// 行合并单元格
	rightLetter := tool.ExcelIndexLetter(index + header.ChildrenLen(val.Children) - 1)
	right, err := excelize.JoinCellName(rightLetter, line)
	if err != nil {
		return fmt.Errorf("left right val join cell name 2 failed. val: %+v err: %+v", val, err)
	}
	if err = er.file.MergeCell(er.NowSheetName, left, right); err != nil {
		return fmt.Errorf("left right val merge cell failed. val: %+v err: %+v", val, err)
	}
	// 为标题行设置样式
	if err = er.file.SetCellStyle(er.NowSheetName, left, right, er.styleId); err != nil {
		return fmt.Errorf("left right val set cell style failed. val: %+v err: %+v", val, err)
	}
	if val.StyleId > 0 {
		if err = er.file.SetCellStyle(er.NowSheetName, left, right, val.StyleId); err != nil {
			return fmt.Errorf("up down val set cell style failed. val: %+v err: %+v", val, err)
		}
	}
	return nil
}

// 上下合并
func (er *Exporter) up2down(val *header.Header, index int, line int) error {
	letter := tool.ExcelIndexLetter(index)
	up, err := excelize.JoinCellName(letter, line)
	if err != nil {
		return fmt.Errorf("up down val join cell name 1 failed. val: %+v err: %+v", val, err)
	}
	down, err := excelize.JoinCellName(letter, er.maxLevel)
	if err != nil {
		return fmt.Errorf("up down val join cell name 2 failed. val: %+v err: %+v", val, err)
	}
	if err = er.file.SetCellValue(er.NowSheetName, up, val.Title); err != nil {
		return fmt.Errorf("up down val set cell value failed. val: %+v err: %+v", val, err)
	}
	// 合并单元格
	if err = er.file.MergeCell(er.NowSheetName, up, down); err != nil {
		return fmt.Errorf("up down val merge cell failed. val: %+v err: %+v", val, err)
	}
	// 为标题行设置样式
	if err = er.file.SetCellStyle(er.NowSheetName, up, down, er.styleId); err != nil {
		return fmt.Errorf("up down val set cell style failed. val: %+v err: %+v", val, err)
	}
	if val.StyleId > 0 {
		if err = er.file.SetCellStyle(er.NowSheetName, up, down, val.StyleId); err != nil {
			return fmt.Errorf("up down val set cell style failed. val: %+v err: %+v", val, err)
		}
	}
	// 批注
	if val.Comment != "" {
		if err = er.file.AddComment(er.NowSheetName, excelize.Comment{
			Cell: fmt.Sprintf("%s%d", letter, line),
			Paragraph: []excelize.RichTextRun{
				{Text: val.Comment},
			},
		}); err != nil {
			return fmt.Errorf("AddComment failed. val: %+v err: %+v", val, err)
		}
	}
	return nil
}

// 默认样式
func (er *Exporter) defaultStyle() *Exporter {
	er.SetStyle(header.GetExportDefaultStyle())
	return er
}

// SetRowStartLine 设置数据开始赋值行号
func (er *Exporter) SetRowStartLine(index int) *Exporter {
	er.startColumn = index
	er.maxLevel = index
	return er
}

// GetSheetIndex 获取excel的第一个sheet索引 （-1=不存在）
func GetSheetIndex(file *excelize.File, sheetName string) int {
	index := -1
	for i, s := range file.GetSheetList() {
		if s == sheetName {
			index = i
			break
		}
	}
	return index
}
