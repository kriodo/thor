package header

import (
	"fmt"
	"github.com/kriodo/thor/office/tool"
	"sort"
	"strings"
)

// Header excel表头参数
type Header struct {
	Title    string     `json:"title,omitempty"`     // [必填]名称
	FieldKey string     `json:"field_key,omitempty"` // [必填]字段值key(必须唯一)
	Children []*Header  `json:"children,omitempty"`  // [非必填]子集
	Id       int64      `json:"id,omitempty"`        // [非必填]id
	Pid      int64      `json:"pid,omitempty"`       // [非必填]pid
	Pkey     string     `json:"pkey,omitempty"`      // [非必填]pkey
	Weight   int        `json:"weight,omitempty"`    // [非必填]排序权重: 越大越靠前
	Export   ExportAttr `json:"export,omitempty"`    // [非必填]导出扩展内容
	Import   ImportAttr `json:"import,omitempty"`    // [非必填]导入扩展内容

	isLast         bool // [内置属性]是否最后一层表头
	level          uint // [内置属性]层级
	childrenMaxLen uint // [内置属性]子集最大数

}

type ExportAttr struct {
	Comment string `json:"comment,omitempty"`  // [非必填]批注
	StyleId int    `json:"style_id,omitempty"` // [非必填]单列的样式ID
}

type ImportAttr struct {
	OtherTitle []string `json:"other_title,omitempty"` // [导入]其他名称：如果title匹配不上会根据此根据集合进行匹配(注意：不验证表头参数)
	LikeTitle  string   `json:"like_title,omitempty"`  // [导入]模糊名称(只支持一级表头)：如果title匹配不上会根据此根据集合进行匹配(注意：不验证表头参数)
	MustExi    bool     `json:"must_exi,omitempty"`    // [导入]验证是否存在
}

// FieldInfo 表头字段的信息
type FieldInfo struct {
	Key       string  `json:"key,omitempty"`        // 字段key
	XIndex    uint    `json:"x_index,omitempty"`    // 字段所在X位置(从0开始)
	YIndex    uint    `json:"y_index,omitempty"`    // 字段所在Y位置(从1开始)
	LastLevel bool    `json:"last_level,omitempty"` // 是否最后一级字段
	MustExi   bool    `json:"must_exi,omitempty"`   // 是否必填
	Width     float64 `json:"width,omitempty"`      // 字段宽度
}

func (h *Header) SetChildrenMaxLen(val uint) {
	h.childrenMaxLen = val
}
func (h *Header) GetChildrenMaxLen() uint {
	return h.childrenMaxLen
}

// GetLevel 获取表头的层级数
func (h *Header) GetLevel() uint {
	return h.level
}

// GetIsLast 获取表头是否是最后一层
func (h *Header) GetIsLast() bool {
	return h.isLast
}

// FormatHeaderInfo 格式化表头tree数据，获取相关表头的相关信息
func FormatHeaderInfo(tree []*Header, level uint, fieldInfo []*FieldInfo) ([]*FieldInfo, error) {
	// 按照 weight 的逆序排序
	sort.Sort(HeadSlice(tree))
	for i, header := range tree {
		childLen := len(header.Children)
		tree[i].level = level
		if childLen == 0 {
			tree[i].isLast = true
			fieldInfo = append(fieldInfo, &FieldInfo{
				Key:       header.FieldKey,
				YIndex:    level,
				LastLevel: childLen == 0,
				MustExi:   header.Import.MustExi,
				Width:     float64(tool.LenChar(header.Title))*1.2 + 8,
			})
			continue
		}
		var err error
		fieldInfo, err = FormatHeaderInfo(header.Children, level+1, fieldInfo)
		if err != nil {
			return nil, err
		}
	}
	return fieldInfo, nil
}

// ListToTree 将list转为tree结构
func ListToTree(list []*Header, pid int64) []*Header {
	var titles []*Header
	for _, item := range list {
		if item.Pid == pid {
			child := ListToTree(list, item.Id)
			subData := item
			subData.Children = child
			titles = append(titles, subData)
		}
	}
	return titles
}

// ChildrenLen 获取子表头长度
func ChildrenLen(children []*Header) uint {
	if len(children) == 0 {
		return 1
	}
	var maxLen uint
	for _, child := range children {
		if len(child.Children) > 0 {
			maxLen += ChildrenLen(child.Children)
		} else {
			maxLen += 1
		}
	}
	return maxLen
}

// MaxLevel 获取最大层级数
func MaxLevel(tree []*Header, level uint) uint {
	for _, v := range tree {
		if v.level > level {
			level = v.level
		}
		if len(v.Children) > 0 {
			level = MaxLevel(v.Children, level)
		}
	}
	return level
}

// FormatTree2LastNameMapV1 获取表头最大数
func FormatTree2LastNameMapV1(tree []*Header) map[string]string {
	dataMap := make(map[string]string, len(tree))
	for _, header := range tree {
		if len(header.Children) == 0 && header.Import.LikeTitle != "" {
			dataMap[header.FieldKey] = header.Import.LikeTitle
		}
	}
	return dataMap
}

func FormatTree2LastNameMapV3(tree []*Header, title string, dataMap map[string]string) map[string]string {
	if !strings.HasPrefix(title, ".") {
		title = SplitTitle(title)
	}
	for _, header := range tree {
		headerTitle := SplitTitle(header.Title)
		newTitle := title + headerTitle
		if len(header.Children) > 0 {
			dataMap = FormatTree2LastNameMapV3(header.Children, newTitle, dataMap)
		} else {
			dataMap[header.FieldKey] = newTitle
		}
	}
	return dataMap
}

// FormatTree2LastNameMapV2 获取表头 [组合中文名称][字段名称]
func FormatTree2LastNameMapV2(tree []*Header, title string, dataMap map[string][]string) map[string][]string {
	if !strings.HasPrefix(title, ".") {
		title = SplitTitle(title)
	}
	for _, header := range tree {
		headerTitle := SplitTitle(header.Title)
		newTitle := title + headerTitle
		if len(header.Children) > 0 {
			dataMap = FormatTree2LastNameMapV2(header.Children, newTitle, dataMap)
			for _, ot := range header.Import.OtherTitle {
				otName := title + SplitTitle(ot)
				dataMap = FormatTree2LastNameMapV2(header.Children, otName, dataMap)
			}
		} else {
			for _, ot := range header.Import.OtherTitle {
				otName := title + SplitTitle(ot)
				dataMap[otName] = append(dataMap[otName], header.FieldKey)
			}
			dataMap[newTitle] = append(dataMap[newTitle], header.FieldKey)
		}
	}
	return dataMap
}

// CheckHeaderId 验证表头的id和pid
func CheckHeaderId(headers []*Header) error {
	validationFieldKey := make(map[string]int)
	for _, header := range headers {
		if header.Id == header.Pid {
			return fmt.Errorf("字段的ID和PID重复: %s", header.FieldKey)
		}
		validationFieldKey[header.FieldKey] += 1
	}
	for key, count := range validationFieldKey {
		if count > 1 {
			return fmt.Errorf("字段FieldKey重复: %s", key)
		}
	}
	return nil
}

// ClearTitle 清除表头里的空格和换行
func ClearTitle(title string) string {
	title = strings.TrimSpace(title)
	title = strings.ReplaceAll(title, " ", "")
	title = strings.ReplaceAll(title, "\n", "")
	return title
}

// SplitTitle 拼接中文标题
func SplitTitle(title string) string {
	if title != "" {
		return "." + title
	}
	return ""
}

type HeadSlice []*Header

func (a HeadSlice) Len() int { // 重写 Len() 方法
	return len(a)
}
func (a HeadSlice) Swap(i, j int) { // 重写 Swap() 方法
	a[i], a[j] = a[j], a[i]
}
func (a HeadSlice) Less(i, j int) bool { // 重写 Less() 方法， 从大到小排序
	return a[j].Weight < a[i].Weight
}
