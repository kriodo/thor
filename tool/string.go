package tool

import (
	"regexp"
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
