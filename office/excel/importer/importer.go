package importer

import (
	"errors"
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
		opts:      nil,
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
			sheetName:       sheetName,
			fieldMap:        make(map[string]*header.FieldInfo),
			letter2TitleMap: make(map[string]*ImportHeaderInfo),
			index2keyMap:    make(map[int]string),
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
	return ir.cur.data
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
		for k := range ir.cur.letter2TitleMap {
			letters = append(letters, k)
		}
		letters = tool.UniqueString(letters)
		tool.SortStrings(letters)
		for _, letter := range letters {
			info, exi := ir.cur.letter2TitleMap[letter]
			if exi {
				val += fmt.Sprintf("%s : %s \n", info.Title, data[info.Title])
			}
		}
	} else {
		for _, info := range ir.cur.fields {
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
		ir.err = fmt.Errorf("获取原始数据错误：%s-%+rows", ir.cur.sheetName, err)
		return ir.err
	}
	sheet.origRows = origRows
	ir.cur.origRows = origRows
	var (
		origLen            = len(origRows)                        // 数据总量
		firstLevelLen      = len(ir.cur.headerTree)               // 第一层表头标题数量
		firstLevelTitleMap = make(map[string]bool, firstLevelLen) // 第一层表头标题
		firstMatchNum      int                                    // 第一层标题匹配上数量
		wantMatchNum       = 2                                    // 需要匹配上的数量
		matchMaxLine       = 50                                   // 匹配表头最大行数，该数据内未匹配到数据就放弃
		isEndMatch         bool                                   // 是否结束匹配
	)
	for _, info := range ir.cur.headerTree {
		firstLevelTitleMap[info.Title] = true
	}
	// 如果表头第一层就一个数据那就1个匹配上即可
	if firstLevelLen == 1 {
		wantMatchNum = 1
	}
	for index, rows := range origRows {
		dataLen := len(rows)
		if ir.cur.maxOrigLen < dataLen {
			ir.cur.maxOrigLen = dataLen
		}
		if isEndMatch {
			continue
		}
		// 超过最大匹配次数，放弃匹配
		if index > matchMaxLine {
			isEndMatch = true
			continue
		}
		for _, row := range rows {
			if firstLevelTitleMap[row] {
				firstMatchNum += 1
			}
		}
		// 达到匹配数量，结束匹配
		if firstMatchNum >= wantMatchNum {
			ir.cur.headerStartLine = index
			isEndMatch = true
		}
	}
	if firstMatchNum < wantMatchNum {
		ir.err = fmt.Errorf("%s 首行表头标题已匹配%d个，请检查表头数据是否正确", ir.cur.sheetName, firstMatchNum)
		return ir.err
	}
	ir.cur.rowStartLine = ir.cur.headerStartLine + int(ir.cur.headerLength)
	if origLen < ir.cur.rowStartLine {
		return nil
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
	ir.cur.letter2TitleMap = make(map[string]*ImportHeaderInfo, ir.cur.maxOrigLen)
	for index, titles := range headerIndex2valMap {
		letter := tool.IndexToLetter(uint(index))
		ir.cur.letter2TitleMap[letter] = &ImportHeaderInfo{
			Index:  index,
			Letter: letter,
			Title:  header.SplitTitleV2(titles),
		}
	}
	// join_title -> 字段key
	title2keyMap := make(map[string]string)
	for _, info := range ir.cur.fields {
		key := header.SplitTitleV2(info.JoinTitle)
		title2keyMap[key] = info.Key
	}
	// 索引对应字段key，未匹配到=join_title
	for i := 0; i < ir.cur.maxOrigLen; i++ {
		letter := tool.IndexToLetter(uint(i))
		fieldInfo, ok := ir.cur.letter2TitleMap[letter]
		if ok && fieldInfo.Title != "" {
			fieldKey, ok2 := title2keyMap[fieldInfo.Title]
			if ok2 {
				ir.cur.index2keyMap[i] = fieldKey
			} else {
				ir.cur.index2keyMap[i] = fieldInfo.Title
			}
		}
	}
	for _, info := range ir.cur.fields {
		// 需要验证表头字段是否存在
		if info.IsRequired {
			joinTitle := header.SplitTitleV2(info.JoinTitle)
			fieldKey, exi := title2keyMap[joinTitle]
			if exi && fieldKey != "" { // 需要验证,并且验证不通过
				// TODO 模糊匹配
				ir.errList = append(ir.errList, fmt.Sprintf("%s 表头名称 %s 不存在", ir.cur.sheetName, joinTitle))
			}
		}
	}
	if len(ir.errList) > 0 {
		ir.err = errors.New(ir.errList[0])
		return ir.err
	}

	return nil
}

// 绑定字段数据
func (ir *Importer) bindImportField() error {
	dataRows := ir.cur.origRows[ir.cur.rowStartLine:]
	for _, rows := range dataRows {
		dataMap := make(map[string]string, len(rows))
		for index, row := range rows {
			if key, exi := ir.cur.index2keyMap[index]; exi {
				dataMap[key] = row
			}
		}
		ir.cur.data = append(ir.cur.data, dataMap)
	}
	return nil
}
