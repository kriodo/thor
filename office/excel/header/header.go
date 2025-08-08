package header

import (
	"fmt"
	"sort"
	"strings"
)

// Header excel表头参数
type Header struct {
	Id             int64     `json:"id,omitempty"`                // id
	Pid            int64     `json:"pid,omitempty"`               // pid
	Pkey           string    `json:"pkey,omitempty"`              // pkey
	Title          string    `json:"title,omitempty"`             // 名称
	Comment        string    `json:"comment,omitempty"`           // 批注
	FieldKey       string    `json:"field_key,omitempty"`         // 字段值key
	Children       []*Header `json:"children,omitempty"`          // 子集
	OtherTitle     []string  `json:"other_title,omitempty"`       // [导入]其他名称：如果title匹配不上会根据此根据集合进行匹配(注意：不验证表头参数)
	LikeTitle      string    `json:"like_title,omitempty"`        // [导入]模糊名称(只支持一级表头)：如果title匹配不上会根据此根据集合进行匹配(注意：不验证表头参数)
	MustExi        bool      `json:"must_exi,omitempty"`          // [导入]验证是否存在
	Weight         int       `json:"weight,omitempty"`            // [导出]排序权重
	StyleId        int       `json:"style_id,omitempty"`          // [导出]单列的样式ID
	Level          int       `json:"level,omitempty"`             // 层级
	IsOnlyOneLevel bool      `json:"is_only_one_level,omitempty"` // 是否只有一个
	IsLastLevel    bool      `json:"is_last_level,omitempty"`     // 是否最后一层级
}

// FormatTree 格式化表头的tree
func FormatTree(tree []*Header, level int, vMap map[string]struct{}) map[string]struct{} {
	if len(vMap) == 0 {
		vMap = make(map[string]struct{})
	}
	// 按照 weight 的逆序排序
	sort.Sort(HeadSlice(tree))
	for i, header := range tree {
		tree[i].Level = level
		l := len(header.Children)
		if header.MustExi && l == 0 {
			vMap[header.FieldKey] = struct{}{}
		}
		if l > 0 {
			vMap = FormatTree(header.Children, level+1, vMap)
		}
		tree[i].IsLastLevel = l == 0
		if tree[i].IsLastLevel && tree[i].Level == 1 {
			tree[i].IsOnlyOneLevel = true
		}
	}
	return vMap
}

// List2Tree 将列表转为tree结构
func List2Tree(list []*Header, pid int64) []*Header {
	var titles []*Header
	for _, item := range list {
		if item.Pid == pid {
			child := List2Tree(list, item.Id)
			d := &Header{
				Id:         item.Id,
				Pid:        item.Pid,
				Title:      item.Title,
				OtherTitle: item.OtherTitle,
				LikeTitle:  item.LikeTitle,
				FieldKey:   item.FieldKey,
				Weight:     item.Weight,
				Children:   child,
				StyleId:    item.StyleId,
				MustExi:    item.MustExi,
			}
			titles = append(titles, d)
		}
	}
	return titles
}

// ChildrenLen 获取子表头长度
func ChildrenLen(children []*Header) int {
	if len(children) == 0 {
		return 1
	}
	var maxLen int
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
func MaxLevel(tree []*Header, level int) int {
	for _, v := range tree {
		if v.Level > level {
			level = v.Level
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
		if header.IsOnlyOneLevel && header.LikeTitle != "" {
			dataMap[header.FieldKey] = header.LikeTitle
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
			for _, ot := range header.OtherTitle {
				otName := title + SplitTitle(ot)
				dataMap = FormatTree2LastNameMapV2(header.Children, otName, dataMap)
			}
		} else {
			for _, ot := range header.OtherTitle {
				otName := title + SplitTitle(ot)
				dataMap[otName] = append(dataMap[otName], header.FieldKey)
			}
			dataMap[newTitle] = append(dataMap[newTitle], header.FieldKey)
		}
	}
	return dataMap
}

// Validation 验证表头
func Validation(headers []*Header) error {
	validationFieldKey := make(map[string]int)
	for _, header := range headers {
		if header.Id == header.Pid {
			return fmt.Errorf("validation header id eq pid . %s", header.FieldKey)
		}
		validationFieldKey[header.FieldKey] += 1
	}
	for key, count := range validationFieldKey {
		if count > 1 {
			return fmt.Errorf("validation field key again defined . %s", key)
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
