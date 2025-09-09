package importer

import (
	"errors"
	"fmt"
	"github.com/kriodo/thor/office/excel/header"
	"github.com/kriodo/thor/office/tool"
	"github.com/xuri/excelize/v2"
	"strings"
)

// 导入v2

type ImportProcessorV2 struct {
	// 基础数据
	file      *excelize.File // 上传文件file [*必填]
	uniqueID  string         // 结果ID
	sheetName string         // 表格sheet名称

	// 表头参数
	noHeader        bool             // 没有表头：此模式获取的是实际表头对应数据
	tempHeader      []*header.Header // 标准表头
	headerLength    int              // 表头层数
	headerStartLine int              // 表头起始行号:从0开始
	rowStartLine    int              // 数据起始行号
	maxDataColNum   int              // 最大数据列数
	maxTitleNum     int              // 最大字段数据（验证用）
	clearUnMatch    bool             // 清除未匹配的表头值

	// 表头处理数据
	tempHeaderZh2NameMap     map[string][]string // (模板)表头: [zh] -> [name] 姓名 => name
	tempHeaderName2ZhMap     map[string]string   // (模板)表头: [name] -> [zh] name => 姓名 （标准）
	tempHeaderName2LikeZhMap map[string]string   // (模板)表头: [name] -> [模糊zh] name => 姓名 （标准）
	realHeaderIndex2ZhMap    map[int]string      // (实际)表头: [index] -> [zh]
	verifyHeaderMap          map[string]struct{} // 需要验证表头的字段：[name] => struct{}

	// 数据参数
	headerIndex2NameMap map[int][]string    // (绑定后)坐标对应字段: [index]=>[name/real zh]
	rows                []map[string]string // 数据

	// 错误
	err       error    // 错误
	errList   []string // 错误集合
	maxErrNum int      // 最大错误数量（大致数值，未做严格要求）
	maxRowNum int      // 最大数据量

	opts []excelize.Options
}

// NewImportProcessor new一个导入处理器
func NewImportProcessor(file *excelize.File) *ImportProcessorV2 {
	return &ImportProcessorV2{
		file:                     file,
		uniqueID:                 tool.GetUUID(),
		sheetName:                "",
		tempHeader:               nil,
		tempHeaderZh2NameMap:     make(map[string][]string),
		tempHeaderName2ZhMap:     make(map[string]string),
		realHeaderIndex2ZhMap:    make(map[int]string),
		tempHeaderName2LikeZhMap: make(map[string]string),
		headerIndex2NameMap:      make(map[int][]string),
		headerLength:             1,
		verifyHeaderMap:          make(map[string]struct{}),
		rowStartLine:             2,
		rows:                     nil,
		err:                      nil,
		maxErrNum:                500,
		maxRowNum:                50000,
		clearUnMatch:             true,
	}
}

func (ipv2 *ImportProcessorV2) clear() {
	ipv2.maxDataColNum = 0
	ipv2.tempHeader = nil
	ipv2.headerLength = 0
	ipv2.headerStartLine = 0
	ipv2.rowStartLine = 0
	ipv2.maxDataColNum = 0
	ipv2.maxTitleNum = 0
	ipv2.tempHeaderZh2NameMap = make(map[string][]string)
	ipv2.tempHeaderName2ZhMap = make(map[string]string)
	ipv2.tempHeaderName2LikeZhMap = make(map[string]string)
	ipv2.realHeaderIndex2ZhMap = make(map[int]string)
	ipv2.verifyHeaderMap = make(map[string]struct{})
	ipv2.headerIndex2NameMap = make(map[int][]string)
	ipv2.rows = []map[string]string{}
	ipv2.err = nil
	ipv2.errList = []string{}
}

// SetTree 设置表头 tree结构
func (ipv2 *ImportProcessorV2) SetTree(tree []*header.Header) *ImportProcessorV2 {
	if len(tree) == 0 {
		ipv2.err = fmt.Errorf("表头不能设置为空")
	}
	ipv2.clear()
	//ipv2.verifyHeaderMap = header.FormatTree(tree, 1, nil)
	ipv2.tempHeader = tree
	//ipv2.headerLength = header.MaxLevel(tree, 1)
	ipv2.rowStartLine = ipv2.headerLength + 1
	return ipv2
}

// SetList 设置list数据 [此模式id、pid必须填写]
func (ipv2 *ImportProcessorV2) SetList(headers []*header.Header) *ImportProcessorV2 {
	ipv2.clear()
	err := header.CheckHeaderId(headers)
	if err != nil {
		ipv2.err = err
		return ipv2
	}
	tree := header.ListToTree(headers, 0)
	//ipv2.verifyHeaderMap = header.FormatTree(tree, 1, nil)
	ipv2.tempHeader = tree
	//ipv2.headerLength = header.MaxLevel(tree, 1)
	ipv2.rowStartLine = ipv2.headerLength + 1
	return ipv2
}

// SetListV2 设置list数据 [此模式pkey必须填写]
func (ipv2 *ImportProcessorV2) SetListV2(headers []*header.Header) *ImportProcessorV2 {
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

	return ipv2.SetList(newHeaders)
}

// SetNoHeader 不设置表头 [此模式下获取数据直接用中文名称获取数据：.姓名 => 张三 ] (注意startLine参数最小值为1)
func (ipv2 *ImportProcessorV2) SetNoHeader(startLine, endLine int) *ImportProcessorV2 {
	if startLine < 1 || endLine < 1 || endLine < startLine {
		ipv2.err = errors.New("表头开始或结束行号参数错误")
		return ipv2
	}
	startLine = startLine - 1

	ipv2.noHeader = true
	ipv2.rowStartLine = endLine + 1
	ipv2.headerStartLine = startLine
	ipv2.headerLength = endLine - startLine
	ipv2.clearUnMatch = false
	return ipv2
}

func (ipv2 *ImportProcessorV2) SetSheetName(sheetName string) *ImportProcessorV2 {
	ipv2.sheetName = sheetName
	return ipv2
}

// SetStartLine 设置表头和数据的其实行号(必须在设置表头之后配置)
func (ipv2 *ImportProcessorV2) SetStartLine(startHeaderLine int, startRowLine int) *ImportProcessorV2 {
	ipv2.headerStartLine = startHeaderLine
	ipv2.rowStartLine = startRowLine
	return ipv2
}

// SetClearUnMatch 配置是否清除未匹配的表头数据
func (ipv2 *ImportProcessorV2) SetClearUnMatch(cum bool) *ImportProcessorV2 {
	ipv2.clearUnMatch = cum
	return ipv2
}

func (ipv2 *ImportProcessorV2) Run() error {
	if ipv2.err != nil {
		return ipv2.err
	}
	rowData, headerData, err := ipv2.handleExcelData()
	if err != nil {
		return err
	}
	if len(headerData) == 0 {
		return errors.New("表格表头为空")
	}
	if len(rowData) == 0 {
		return errors.New("表格数据为空")
	}
	if !ipv2.noHeader {
		// 处理模板表头
		err = ipv2.handleTempHeaderMap()
		if err != nil {
			return err
		}
	}
	// 绑定数据到表头
	ipv2.getRealHeaderMap(headerData)
	// 验证表头&绑定表头坐标
	err = ipv2.vTempHeaderAndBindIndex()
	if err != nil {
		return err
	}
	// 绑定数据->行数据
	ipv2.bindData2RowData(rowData)

	return nil
}

// GetRows 获取值
func (ipv2 *ImportProcessorV2) GetRows() []map[string]string {
	return ipv2.rows
}

// GetRowStartLine 数据开始行号
func (ipv2 *ImportProcessorV2) GetRowStartLine() int {
	return ipv2.rowStartLine
}

// GetMaxDataColNum 获取最大数据列数
func (ipv2 *ImportProcessorV2) GetMaxDataColNum() int {
	return ipv2.maxDataColNum
}

// GetTitleByKey 根据key获取中文描述
func (ipv2 *ImportProcessorV2) GetTitleByKey(key string) []string {
	return ipv2.tempHeaderZh2NameMap[key]
}

// GetErrList 获取多个错误
func (ipv2 *ImportProcessorV2) GetErrList() []string {
	return ipv2.errList
}

func (ipv2 *ImportProcessorV2) handleTempHeaderMap() error {
	ipv2.tempHeaderName2LikeZhMap = header.FormatTree2LastNameMapV1(ipv2.tempHeader)
	dataMap1 := make(map[string][]string)
	ipv2.tempHeaderZh2NameMap = header.FormatTree2LastNameMapV2(ipv2.tempHeader, "", dataMap1)
	dataMap2 := make(map[string]string)
	ipv2.tempHeaderName2ZhMap = header.FormatTree2LastNameMapV3(ipv2.tempHeader, "", dataMap2)
	return nil
}

// 验证表头&绑定表头坐标
func (ipv2 *ImportProcessorV2) vTempHeaderAndBindIndex() error {
	var errList []string
	realHeaderName2IndexMap := make(map[string]int)
	for index, zh := range ipv2.realHeaderIndex2ZhMap { // 实际数据
		// 绑定坐标（重要）
		name, exi := ipv2.tempHeaderZh2NameMap[zh]
		if !exi {
			if !ipv2.clearUnMatch {
				realHeaderName2IndexMap[zh] = index
				ipv2.headerIndex2NameMap[index] = append(ipv2.headerIndex2NameMap[index], zh)
			}
			continue
		}
		ipv2.headerIndex2NameMap[index] = append(ipv2.headerIndex2NameMap[index], name...)
		for _, nm := range name {
			realHeaderName2IndexMap[nm] = index
		}
	}
	// 利用模糊再查一下
	for name, likeZh := range ipv2.tempHeaderName2LikeZhMap {
		if index, _ := realHeaderName2IndexMap[name]; index > 0 {
			continue
		}
		for index, zh := range ipv2.realHeaderIndex2ZhMap { // 实际数据
			if strings.Contains(zh, likeZh) {
				ipv2.headerIndex2NameMap[index] = append(ipv2.headerIndex2NameMap[index], name)
				realHeaderName2IndexMap[name] = index
			}
		}

	}
	for name, zh := range ipv2.tempHeaderName2ZhMap {
		_, ok := ipv2.verifyHeaderMap[name]
		index, _ := realHeaderName2IndexMap[name]
		if ok && index == 0 { // 需要验证,并且验证不通过
			errList = append(errList, fmt.Sprintf("%s不存在", showTitle(zh)))
		}
	}
	tool.UniqueString(errList)
	tool.SortStrings(errList)
	ipv2.errList = errList

	if len(errList) > 0 {
		return errors.New(errList[0])
	}

	return nil
}

// 绑定数据
func (ipv2 *ImportProcessorV2) bindData2RowData(rowData [][]string) {
	var rows []map[string]string
	for _, row := range rowData {
		dataMap := make(map[string]string, len(row))
		for col, val := range row {
			index := col + 1
			if name, exi := ipv2.headerIndex2NameMap[index]; exi {
				for _, vn := range name {
					dataMap[vn] = val
				}
			}
		}
		rows = append(rows, dataMap)
	}
	ipv2.rows = rows
}

// 获取excel实际表头数据
func (ipv2 *ImportProcessorV2) getRealHeaderMap(headerData [][]string) {
	var (
		maxCol int
		colMap = make(map[int][]string) // [列号][名称]
	)
	for _, v := range headerData {
		length := len(v)
		if maxCol < length {
			maxCol = length
		}
	}

	lastOriginValMap := make(map[int]string, maxCol)
	for j, v1 := range headerData {
		rowLen := len(v1)
		if rowLen == 0 {
			continue
		}
		var lastVal string
		for i := 1; i <= maxCol; i++ {
			if lastOriginValMap[i] != "" {
				lastVal = ""
			}
			if i > rowLen {
				colMap[i] = append(colMap[i], lastVal)
			} else {
				val := header.ClearTitle(v1[i-1])
				lastOriginValMap[i] = val
				if val == "" && j != len(headerData)-1 {
					val = lastVal
				}
				colMap[i] = append(colMap[i], val)
				lastVal = val
			}
		}
	}
	newColMap := make(map[int]string)
	for col, names := range colMap {
		var str string
		for _, name := range names {
			if name == "" {
				continue
			}
			str = str + header.SplitTitle(name)
		}
		newColMap[col] = str
	}
	// todo debug打印
	//var slc []string
	//for i := 1; i <= maxCol; i++ {
	//	slc = append(slc, newColMap[i])
	//}
	//str := xdutil.Json(slc)
	//_ = str
	ipv2.realHeaderIndex2ZhMap = newColMap
}

func (ipv2 *ImportProcessorV2) SetOpts(opts ...excelize.Options) *ImportProcessorV2 {
	ipv2.opts = opts
	return ipv2
}

// 获取处理数据的excel数据: 将表头和数据进行分离
func (ipv2 *ImportProcessorV2) handleExcelData() ([][]string, [][]string, error) {
	if ipv2.sheetName == "" {
		sheetNameList := ipv2.file.GetSheetList()
		if len(sheetNameList) > 0 {
			ipv2.sheetName = sheetNameList[0]
		}
	}
	// 获取处理excel数据
	rows, err := tool.ReadExcel2Opt(ipv2.file, ipv2.sheetName, ipv2.opts...)
	if err != nil {
		return nil, nil, err
	}
	length := len(rows) // 所有数据长度
	if length <= 0 {
		return nil, nil, errors.New("文件有效数据为空")
	}
	if ipv2.headerLength > length {
		return nil, nil, errors.New("表格表头错误，请使用正确的表格")
	}
	// excel数据行数限制
	if length > ipv2.maxRowNum {
		return nil, nil, fmt.Errorf("文件数据量超限：%d", ipv2.maxErrNum)
	}
	if ipv2.rowStartLine > len(rows) {
		return nil, nil, fmt.Errorf("数据起始行号错误：%d", ipv2.rowStartLine)
	}
	if ipv2.headerStartLine > len(rows) {
		return nil, nil, fmt.Errorf("表头起始行号错误：%d", ipv2.headerStartLine)
	}

	for _, row := range rows {
		rowLen := len(row)
		if rowLen > ipv2.maxDataColNum {
			ipv2.maxDataColNum = rowLen
		}
	}

	var i int
	for k, v := range rows[ipv2.rowStartLine:] {
		if len(v) != 0 {
			break
		}
		i = k + 1
	}

	splitIndex := ipv2.headerStartLine + ipv2.headerLength + i
	var (
		header [][]string
		data   [][]string
	)
	if length >= splitIndex {
		header = rows[ipv2.headerStartLine:splitIndex]
		ipv2.rowStartLine += i
		data = rows[splitIndex:]
	}

	return data, header, nil
}

func showTitle(title string) string {
	arr := strings.Split(title, ".")
	return strings.TrimSpace(strings.Join(arr, " "))
}

func (ipv2 *ImportProcessorV2) GetRealHeaderIndex2ZhMap() map[int]string {
	return ipv2.realHeaderIndex2ZhMap
}
