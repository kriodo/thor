package importer

import (
	"fmt"
	"github.com/kriodo/thor/office/excel/header"
	"github.com/kriodo/thor/office/tool"
	"github.com/xuri/excelize/v2"
)

// Importer 导入
type Importer struct {
	// 基础数据
	file     *excelize.File          // 上传文件file [*必填]
	cur      *ImportSheet            // 当前sheet
	sheet    map[string]*ImportSheet // sheet数据
	uniqueID string                  // 结果ID

	// 错误
	err       error              // 错误
	errList   []string           // 错误集合
	maxErrNum int                // 最大错误数量（大致数值，未做严格要求）
	maxRowNum int                // 最大数据量
	opts      []excelize.Options // 自定义的原始配置信息
}

func NewImporter(file *excelize.File) *Importer {
	ir := &Importer{
		file:      file,
		cur:       nil,
		sheet:     nil,
		uniqueID:  tool.GetUUID(),
		err:       nil,
		errList:   nil,
		maxErrNum: 0,
		maxRowNum: 0,
	}
	if ir.file == nil {
		ir.err = fmt.Errorf("导入文件不能为空")
		return ir
	}
	// 获取excel的sheet数据
	sheetList := file.GetSheetList()
	sheetLen := len(sheetList)
	if sheetLen == 0 {
		ir.err = fmt.Errorf("导入文件sheet无数据")
		return ir
	}
	ir.sheet = make(map[string]*ImportSheet, sheetLen)
	for i, sheetName := range sheetList {
		sheet := &ImportSheet{
			sheetName:    sheetName,
			fieldInfoMap: make(map[string]*header.FieldInfo),
		}
		if i == 0 {
			ir.cur = sheet
		}
		ir.sheet[sheetName] = sheet
	}

	return ir
}

// CutSheet 切换sheet
func (ir *Importer) CutSheet(sheetName string) error {
	sheet, exi := ir.sheet[sheetName]
	if !exi {
		ir.err = fmt.Errorf("sheet不存在：%s", sheetName)
		return ir.err
	}
	ir.cur = sheet
	return nil
}

// SetOpts 设置自定义的原始配置信息
func (ir *Importer) SetOpts(opts []excelize.Options) error {
	ir.opts = opts
	return nil
}

// GetRows 获取数据
func (ir *Importer) GetRows() []map[string]string {
	if !ir.cur.isSet {
		ir.err = fmt.Errorf("未设置表头信息：%s", ir.cur.sheetName)
		return nil
	}
	ir.err = ir.bindImportField()
	if ir.err != nil {
		return nil
	}
	sheet, exi := ir.sheet[ir.cur.sheetName]
	if !exi {
		ir.err = fmt.Errorf("sheet不存在：%s", ir.cur.sheetName)
		return nil
	}
	ir.cur = sheet
	return ir.cur.field2dataMap
}

// 获取错误
func (ir *Importer) Error() error {
	return ir.err
}

// PrintRows 打印数据结果
func (ir *Importer) PrintRows(data map[string]string) string {
	val := "\n"
	if ir.cur.noHeader {
		var letters []string
		for k := range ir.cur.headerTitleInfos {
			letters = append(letters, k)
		}
		letters = tool.UniqueString(letters)
		tool.SortStrings(letters)
		for _, letter := range letters {
			info, exi := ir.cur.headerTitleInfos[letter]
			if exi {
				val += fmt.Sprintf("%s : %s \n", info.Title, data[info.Title])
			}
		}
	} else {
		for _, info := range ir.cur.fieldInfos {
			val += fmt.Sprintf("%s : %s \n", header.SplitTitleV2(info.JoinTitle), data[info.Key])
		}
	}
	return val
}

// ----------------------------------------[ 分割线 ]-------------------------------------------//

// handleRowBefore 处理前置操作（验证数据、绑定表头、验证表头）
func (ir *Importer) handleRowBefore() error {
	sheet, exi := ir.sheet[ir.cur.sheetName]
	if !exi {
		ir.err = fmt.Errorf("sheet不存在：%s", ir.cur.sheetName)
		return ir.err
	}
	origRows, err := ir.getOrigRows()
	if err != nil {
		ir.err = fmt.Errorf("获取原始数据错误：%s-%+v", ir.cur.sheetName, err)
		return ir.err
	}
	sheet.origRows = origRows
	ir.cur.origRows = origRows
	origLen := len(origRows)
	var hasData bool
	for i, v := range origRows {
		dataLen := len(v)
		if dataLen > 0 && !hasData {
			ir.cur.headerStartLine = i
			hasData = true
		}
		if ir.cur.maxOrigLen < dataLen {
			ir.cur.maxOrigLen = dataLen
		}
	}
	ir.cur.rowStartLine = ir.cur.headerStartLine + int(ir.cur.headerLength)
	if origLen < ir.cur.rowStartLine {
		return nil
	}
	ir.err = ir.bindImportHeader()
	if ir.err != nil {
		return ir.err
	}
	return nil
}

// 获取sheet所有原始单元格数据
func (ir *Importer) getOrigRows() ([][]string, error) {
	if len(ir.cur.origRows) > 0 {
		return ir.cur.origRows, nil
	}
	rows, err := ir.file.GetRows(ir.cur.sheetName, ir.opts...)
	if err != nil {
		ir.err = fmt.Errorf("读取表格失败: %s-%+v", ir.cur.sheetName, err)
		return nil, ir.err
	}

	return rows, nil
}

// 绑定表头数据
func (ir *Importer) bindImportHeader() error {
	var (
		headerOrigRows       = ir.cur.origRows[ir.cur.headerStartLine:ir.cur.rowStartLine]
		headerOrigRowLen     = len(headerOrigRows)
		headerIndex2valMap   = make(map[int][]string, ir.cur.maxOrigLen) // 表头字段索引对应数据map
		headerLastOrigValMap = make(map[int]string, ir.cur.maxOrigLen)
	)
	// 合并单元格的数据会出现空数据情况
	for i, rows := range headerOrigRows {
		rowLen := len(rows)
		if rowLen == 0 {
			continue
		}
		var lastVal string
		for ii := 0; ii < ir.cur.maxOrigLen; ii++ {
			if headerLastOrigValMap[ii] != "" {
				lastVal = ""
			}
			if ii >= rowLen {
				headerIndex2valMap[ii] = append(headerIndex2valMap[ii], lastVal)
			} else {
				val := header.ClearTitle(rows[ii])
				headerLastOrigValMap[ii] = val
				if val == "" && i != headerOrigRowLen {
					val = lastVal
				}
				headerIndex2valMap[ii] = append(headerIndex2valMap[ii], val)
				lastVal = val
			}
		}
	}
	ir.cur.headerTitleInfos = make(map[string]*ImportHeaderInfo, ir.cur.maxOrigLen)
	for index, titles := range headerIndex2valMap {
		letter := tool.IndexToLetter(uint(index))
		ir.cur.headerTitleInfos[letter] = &ImportHeaderInfo{
			Index:  index,
			Letter: letter,
			Title:  header.SplitTitleV2(titles),
		}
	}
	return nil
}

// 绑定字段数据
func (ir *Importer) bindImportField() error {
	dataRows := ir.cur.origRows[ir.cur.rowStartLine:]
	title2keyMap := make(map[string]string)
	for _, info := range ir.cur.fieldInfos {
		key := header.SplitTitleV2(info.JoinTitle)
		title2keyMap[key] = info.Key
	}
	for _, rows := range dataRows {
		dataMap := make(map[string]string, len(rows))
		for index, row := range rows {
			letter := tool.IndexToLetter(uint(index))
			fieldInfo, ok := ir.cur.headerTitleInfos[letter]
			if ok && fieldInfo.Title != "" {
				fieldKey, ok2 := title2keyMap[fieldInfo.Title]
				if ok2 {
					dataMap[fieldKey] = row
				} else {
					dataMap[fieldInfo.Title] = row
				}
			}
		}
		ir.cur.field2dataMap = append(ir.cur.field2dataMap, dataMap)
	}
	//var errList []string
	//realHeaderName2IndexMap := make(map[string]int)
	//for index, zh := range ir.cur.realHeaderIndex2ZhMap { // 实际数据
	//	// 绑定坐标（重要）
	//	name, exi := ir.cur.tempHeaderZh2NameMap[zh]
	//	if !exi {
	//		if !ir.cur.clearUnMatch {
	//			realHeaderName2IndexMap[zh] = index
	//			ir.cur.headerIndex2NameMap[index] = append(ir.cur.headerIndex2NameMap[index], zh)
	//		}
	//		continue
	//	}
	//	ir.cur.headerIndex2NameMap[index] = append(ir.cur.headerIndex2NameMap[index], name...)
	//	for _, nm := range name {
	//		realHeaderName2IndexMap[nm] = index
	//	}
	//}
	//// 利用模糊再查一下
	//for name, likeZh := range ir.cur.tempHeaderName2LikeZhMap {
	//	if index, _ := realHeaderName2IndexMap[name]; index > 0 {
	//		continue
	//	}
	//	for index, zh := range ir.cur.realHeaderIndex2ZhMap { // 实际数据
	//		if strings.Contains(zh, likeZh) {
	//			ir.cur.headerIndex2NameMap[index] = append(ir.cur.headerIndex2NameMap[index], name)
	//			realHeaderName2IndexMap[name] = index
	//		}
	//	}
	//
	//}
	//for name, zh := range ir.cur.tempHeaderName2ZhMap {
	//	ok := ir.cur.fieldCheckMap[name]
	//	index, _ := realHeaderName2IndexMap[name]
	//	if ok && index == 0 { // 需要验证,并且验证不通过
	//		errList = append(errList, fmt.Sprintf("%s不存在", showTitle(zh)))
	//	}
	//}
	//tool.UniqueString(errList)
	//tool.SortStrings(errList)
	//ir.errList = errList
	//if len(errList) > 0 {
	//	return errors.New(errList[0])
	//}

	return nil
}
