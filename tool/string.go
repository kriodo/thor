package tool

import (
	"regexp"
	"sort"
	"strings"
)

// LenChar 中英文符号字符串长度
func LenChar(s string) int {
	return strings.Count(s, "") - 1
}

// ContainsEnglishWord 判断字符串中是否包含英文单词
func ContainsEnglishWord(s string) bool {
	// 匹配至少一个英文字母组成的单词（支持大小写）
	re := regexp.MustCompile(`[A-Za-z]+`)
	return re.MatchString(s)
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
