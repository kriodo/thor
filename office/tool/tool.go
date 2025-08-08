package tool

import (
	"encoding/json"
	"errors"
	"fmt"
	"github.com/kriodo/thor/office/excel/constant"
	"github.com/kriodo/thor/tool"
	"github.com/segmentio/ksuid"
	"github.com/xuri/excelize/v2"
	"log"
	"math"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"
)

// IndexToLetter 索引转excel字母列名 0=>A 1=>B 26=>AA
func IndexToLetter(index int) string {
	if index < 0 {
		return ""
	}
	var abc = []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	result := abc[index%26]
	index = index / 26
	for index > 0 {
		index = index - 1
		result = abc[index%26] + result
		index = index / 26
	}
	return result
}

// 读取excel
func readExcel(fileDir, sheetName string) ([][]string, error) {
	var err error
	f, err := excelize.OpenFile(fileDir)
	if err != nil {
		return nil, err
	}
	defer func() {
		if err = f.Close(); err != nil {
			log.Println(err)
		}
	}()

	// 获取sheet所有单元格
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, errors.New("读取表格失败")
	}

	return rows, nil
}

// 读取excel
func readExcel2(f *excelize.File, sheetName string) ([][]string, error) {
	// 获取sheet所有单元格
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, errors.New("读取表格失败")
	}

	return rows, nil
}

// 读取excel
func ReadExcel2Opt(f *excelize.File, sheetName string, opts ...excelize.Options) ([][]string, error) {
	// 获取sheet所有单元格
	rows, err := f.GetRows(sheetName, opts...)
	if err != nil {
		return nil, errors.New("读取表格失败")
	}

	return rows, nil
}

// verifyHeader 验证表头
func verifyHeader(temHeader, header [][]string) error {
	var (
		l1, l2 = len(temHeader), len(header)
		errStr = "表格表头不正确，请使用正确表格表头"
	)
	if l1 != l2 {
		return errors.New(errStr)
	}
	for i := 0; i < l1; i++ {
		l11, l22 := len(temHeader[i]), len(header[i])
		if l11 != l22 {
			return fmt.Errorf("第%d行的%s", i+1, errStr)
		}
		for j := 0; j < l11; j++ {
			if temHeader[i][j] != header[i][j] {
				return fmt.Errorf("第%d行，第%d列%s", i+1, j+1, errStr)
			}
		}
	}
	return nil
}

// stringMatchExport 根据字符串匹配内容
func stringMatchExport(str string, reg *regexp.Regexp) (res string, err error) {
	defer func() {
		if panicInfo := recover(); panicInfo != nil {
			err = errors.New("not match reg")
		}
	}()
	return reg.FindStringSubmatch(str)[1], nil
}

// 获取模板表头
func GetTempHeader(f *excelize.File, sheetName string) ([][]string, error) {
	excel, err := readExcel2(f, sheetName)
	if err != nil {
		return nil, err
	}

	if len(excel) <= 0 {
		return nil, errors.New("模板未设置表头")
	}
	return excel, nil
}

// 列唯一性校验
func (p *eimport.Processor) uniqueFormat(uniqueKey string, col string, mappingField map[string]string) string {
	format, ok := mappingField[ExcelUniqueSign]
	if ok && format == "true" {
		uniqueKey += "_" + col
	}
	return uniqueKey
}

// 格式化时间
func (p *Processor) dateFormat(col *string, mappingField map[string]string) (errList []string) {
	format, ok := mappingField[ExcelDateSign]
	if !ok || format == "" {
		return
	}
	// 初步判断是否是excel日期 44562
	if len(*col) == 5 {
		f, err := strconv.ParseFloat(*col, 64)
		if err != nil {
			return
		}
		t, err := excelize.ExcelDateToTime(f, false)
		if err == nil {
			*col = t.Format(format)
			return
		}
	}
	// 列举所有使用的excel时间格式
	formatSlc1 := []string{"2006/01", "2006-01", "200601", "2006/1", "2006-1", "20061"}              // 年月
	formatSlc2 := []string{"2006/01/02", "2006-01-02", "20060102", "2006/1/2", "2006-1-2", "200612"} // 年月日
	if len(format) <= 7 {
		for _, v := range formatSlc1 {
			location, err := time.ParseInLocation(v, *col, time.Local)
			if err != nil {
				continue
			}
			*col = location.Format(format)
			return
		}
	} else {
		for _, v := range formatSlc2 {
			location, err := time.ParseInLocation(v, *col, time.Local)
			if err != nil {
				continue
			}
			*col = location.Format(format)
			return
		}
	}
	return []string{fmt.Sprintf("%s格式错误", mappingField[ExcelZhTitleSign])}
}

// PageSize 切片分页数，用于切片分批操作
func PageSize(total, groupSize int) int {
	return int(math.Ceil(float64(total) / float64(groupSize))) //page总数
}

// ExcelIndexLetter 根据行的index获取行字母
func ExcelIndexLetter(rowIndex int) string {
	letters := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	result := letters[rowIndex%26]
	rowIndex = rowIndex / 26
	for rowIndex > 0 {
		rowIndex = rowIndex - 1
		result = letters[rowIndex%26] + result
		rowIndex = rowIndex / 26
	}

	return result
}

func Any2String(value interface{}) string {
	if value == nil {
		return ""
	}
	switch value.(type) {
	case float64: // 浮点型 3.0将会转换成字符串"3"
		ft, err := value.(float64)
		if !err {
			// logger
		}
		return strconv.FormatFloat(ft, 'f', -1, 64)
	case float32:
		ft := value.(float32)
		return strconv.FormatFloat(float64(ft), 'f', -1, 64)
	case int:
		it := value.(int)
		return strconv.Itoa(it)
	case uint:
		it := value.(uint)
		return strconv.Itoa(int(it))
	case int8:
		it := value.(int8)
		return strconv.Itoa(int(it))
	case uint8:
		it := value.(uint8)
		return strconv.Itoa(int(it))
	case int16:
		it := value.(int16)
		return strconv.Itoa(int(it))
	case uint16:
		it := value.(uint16)
		return strconv.Itoa(int(it))
	case int32:
		it := value.(int32)
		return strconv.Itoa(int(it))
	case uint32:
		it := value.(uint32)
		return strconv.Itoa(int(it))
	case int64:
		it := value.(int64)
		return strconv.FormatInt(it, 10)
	case uint64:
		it := value.(uint64)
		return strconv.FormatUint(it, 10)
	case string:
		return value.(string)
	case []byte:
		return string(value.([]byte))
	default:
		newValue, _ := json.Marshal(value)
		return string(newValue)
	}
}

type ExcelDropListReq struct {
	XIndex         int      // x轴坐标
	YIndex         int      // y轴坐标
	YMaxIndex      int      // y轴坐标（此值存在就按照最大值处理 默认=YIndex）
	DropList       []string // 下拉数据
	IsMax          bool     // 是否超最大
	DropSheetName  string   // 下拉sheet name
	QuoteSheetName string   // 引用sheet name
	DropKey        string   // 下拉定位key
}

// SetDropList 处理excel的下拉框属性(超过255字符就新建sheet)
func SetDropList(f *excelize.File, req *ExcelDropListReq) (dropKey string, err error) {
	if req.IsMax {
		// 新建一个sheet页
		if req.DropSheetName == "" {
			req.DropSheetName = "下拉数据"
		}
		dropKey = req.DropKey
		if dropKey == "" {
			_, err = f.NewSheet(req.DropSheetName)
			if err != nil {
				return
			}

			// 在新建的sheet页插入数据
			length := len(req.DropList)
			for i := 0; i < length; i++ {
				cell := fmt.Sprintf("A%d", i+1)
				err = f.SetCellValue(req.DropSheetName, cell, req.DropList[i])
				if err != nil {
					return
				}
			}
			// 隐藏新建的sheet页
			err = f.SetSheetVisible(req.DropSheetName, false)
			if err != nil {
				return
			}
			dropKey = fmt.Sprintf("%s!$A$1:$A$%d", req.DropSheetName, length)
		}

		// 创建一个数据验证规则，引用新建sheet页的数据
		dv := excelize.NewDataValidation(true)
		x := IndexToLetter(req.XIndex)
		if req.YMaxIndex <= 0 {
			req.YMaxIndex = req.YIndex
		}
		dv.Sqref = fmt.Sprintf("%s%d:%s%d", x, req.YIndex, x, req.YMaxIndex) // 设置较大的范围
		dv.SetSqrefDropList(dropKey)

		// 在第一个sheet页应用数据验证规则
		err = f.AddDataValidation(req.QuoteSheetName, dv)
		if err != nil {
			return
		}
	} else {
		//if req.YMaxIndex > 0 {
		//	req.YIndex = req.YMaxIndex
		//}
		dv := excelize.NewDataValidation(true)
		x := IndexToLetter(req.XIndex)
		dv.SetSqref(fmt.Sprintf("%s%d", x, req.YIndex))
		err = dv.SetDropList(req.DropList)
		if err != nil {
			return
		}
		dfSheetName := constant.DefaultSheetName
		if req.QuoteSheetName != "" {
			dfSheetName = req.QuoteSheetName
		}
		err = f.AddDataValidation(dfSheetName, dv)
		if err != nil {
			return
		}
	}
	return
}

// FormatSheetName 格式化sheet名称，防止过长报错
func FormatSheetName(sheetName string) string {
	l := tool.LenChar(sheetName)
	if l <= 0 {
		return ""
	}
	if l <= 30 {
		return sheetName
	}
	newStr := []rune(sheetName)
	return string(newStr[:30])
}

// CheckSheetName 验证sheet名称是否符合长度
func CheckSheetName(sheetName string) error {
	charLen := tool.LenChar(sheetName)
	if charLen <= 0 {
		return fmt.Errorf("sheet名称不能为空")
	}
	if charLen > 30 {
		return fmt.Errorf("sheet名称长度过长")
	}
	return nil
}

func FormatFileName(sheetName string) string {
	l := tool.LenChar(sheetName)
	if l <= 0 {
		return ""
	}
	if l <= 200 {
		return sheetName
	}
	newStr := []rune(sheetName)
	return string(newStr[:200])
}

func UniqueString(arr []string) []string {
	dataMap := make(map[string]struct{}, len(arr))
	newArr := make([]string, 0, len(dataMap))
	for _, s := range arr {
		if _, ok := dataMap[s]; ok {
			continue
		}
		newArr = append(newArr, s)
		dataMap[s] = struct{}{}
	}

	return newArr
}

// LenChar 中英文符号字符串长度
func LenChar(s string) int {
	return strings.Count(s, "") - 1
}

func GetUUID() string {
	return ksuid.New().String()
}

type ByString []string

func (s ByString) Len() int {
	return len(s)
}
func (s ByString) Swap(i, j int) {
	s[i], s[j] = s[j], s[i]
}
func (s ByString) Less(i, j int) bool {
	return s[i] < s[j] // 这里定义了升序排序，根据需要可以改为s[i] > s[j]实现降序
}

// SortStrings 对字符串数组进行排序。
// 它接受一个字符串切片作为输入，并通过定义的排序规则对其进行排序。
// 排序是通过嵌套的ByString类型实现的，该类型实现了sort.Interface接口的三个方法，
// 从而允许sort包中的排序算法对字符串切片进行排序。
func SortStrings(arr []string) {
	sort.Sort(ByString(arr))
}

// InStringArray 检测给定的值是包含string切片之中
func InStringArray(i string, arr []string) bool {
	if len(arr) == 0 {
		return false
	}
	for _, v := range arr {
		if i == v {
			return true
		}
	}

	return false
}
