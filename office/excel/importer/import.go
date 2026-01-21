package importer

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"regexp"
	"strconv"
	"strings"
	"unicode"
)

func FastRowCount1(filePath, sheetName string) (int, error) {
	// 1. 用 excelize 打开文件
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return 0, err
	}
	defer f.Close()

	// 2. 获取 dimension 预估
	dim, err := f.GetSheetDimension(sheetName)
	if err != nil || dim == "" {
		return 0, errors.New("无法获取dimension")
	}
	parts := strings.Split(dim, ":")
	if len(parts) != 2 {
		return 0, errors.New("dimension格式错误")
	}

	cell := parts[1] // 右下角单元格，例如 "F412345"
	i := 0
	for ; i < len(cell); i++ {
		if unicode.IsDigit(rune(cell[i])) {
			break
		}
	}
	estimatedRow, _ := strconv.Atoi(cell[i:])

	// 3. 快速尾部扫描 sheet.xml，修正 dimension 偏差
	maxRow, err := scanSheetMaxRow(filePath, sheetName)
	if err != nil {
		return estimatedRow, nil // 尾部扫描失败就用 dimension 估算
	}

	// 返回 dimension 和实际尾部扫描中的较大值
	if maxRow > estimatedRow {
		return maxRow, nil
	}
	return estimatedRow, nil
}

func scanSheetMaxRow(filePath, sheetName string) (int, error) {
	r, err := zip.OpenReader(filePath)
	if err != nil {
		return 0, err
	}
	defer r.Close()

	target := "xl/worksheets/" + sheetName + ".xml"
	var sheetFile *zip.File
	for _, f := range r.File {
		if f.Name == target {
			sheetFile = f
			break
		}
	}
	if sheetFile == nil {
		return 0, errors.New("sheet not found")
	}

	rc, err := sheetFile.Open()
	if err != nil {
		return 0, err
	}
	defer rc.Close()

	decoder := xml.NewDecoder(rc)
	maxRow := 0
	for {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		if se, ok := tok.(xml.StartElement); ok && se.Name.Local == "row" {
			for _, attr := range se.Attr {
				if attr.Name.Local == "r" {
					row, _ := strconv.Atoi(attr.Value)
					if row > maxRow {
						maxRow = row
					}
					break
				}
			}
		}
	}
	return maxRow, nil
}

func FastRowCount2(filePath string, sheet string) (int, error) {
	// 1. 用 excelize 打开文件
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return 0, err
	}
	defer f.Close()
	rows, err := f.Rows(sheet)
	if err != nil {
		return 0, err
	}
	defer rows.Close()

	cnt := 0
	for rows.Next() {
		cnt++
		// ❗ 不取 Cell 值 → 速度极快
		_, _ = rows.Columns() // 如果你不需要列内容，这行也可以删除
	}
	return cnt, nil
}

func FastRowCount3(filePath string, sheetName string) (int, error) {

	// 真实 Sheet 文件名
	xmlName := fmt.Sprintf("xl/worksheets/%s.xml", strings.ToLower(sheetName))

	// 2. 打开 ZIP（xlsx）
	zr, err := zip.OpenReader(filePath)
	if err != nil {
		return 0, err
	}
	defer zr.Close()

	// 3. 找对应的 xml 文件
	var sheetFile *zip.File
	for _, f := range zr.File {
		println()
		if f.Name == xmlName {
			sheetFile = f
			break
		}
	}
	if sheetFile == nil {
		return 0, fmt.Errorf("sheet xml not found: %s", xmlName)
	}

	rc, err := sheetFile.Open()
	if err != nil {
		return 0, err
	}
	defer rc.Close()

	reader := bufio.NewReader(rc)
	maxRow := 0
	pattern := []byte(`r="`)

	// 4. 超快速扫描 XML 找最大 row
	for {
		line, err := reader.ReadBytes('\n')
		if err != nil {
			if err == io.EOF {
				break
			}
			return 0, err
		}

		idx := bytes.Index(line, pattern)
		if idx == -1 {
			continue
		}

		start := idx + len(pattern)
		end := start

		for end < len(line) && line[end] >= '0' && line[end] <= '9' {
			end++
		}

		rowNum, _ := strconv.Atoi(string(line[start:end]))
		if rowNum > maxRow {
			maxRow = rowNum
		}
	}

	return maxRow, nil
}

var rowRe = regexp.MustCompile(`<row[^>]* r="([0-9]+)"`)

func FastRowCount4(xlsxPath, sheetXML string) (int, error) {
	r, err := zip.OpenReader(xlsxPath)
	if err != nil {
		return 0, err
	}
	defer r.Close()

	sheetXML = fmt.Sprintf("xl/worksheets/%s.xml", strings.ToLower(sheetXML))

	for _, f := range r.File {
		if f.Name == sheetXML { // xl/worksheets/sheet1.xml
			rc, _ := f.Open()
			defer rc.Close()

			max := 0
			buf := make([]byte, 1024*1024)
			for {
				n, err := rc.Read(buf)
				if n > 0 {
					matches := rowRe.FindAllSubmatch(buf[:n], -1)
					for _, m := range matches {
						// 转成 int
						var r int
						fmt.Sscanf(string(m[1]), "%d", &r)
						if r > max {
							max = r
						}
					}
				}
				if err != nil {
					break
				}
			}
			return max, nil
		}
	}
	return 0, fmt.Errorf("sheet xml not found")
}

func FastRowCount5(path string, sheetXML string) (int, error) {

	fzip, err := zip.OpenReader(path)
	if err != nil {
		return 0, err
	}
	defer fzip.Close()
	sheetXML = fmt.Sprintf("xl/worksheets/%s.xml", strings.ToLower(sheetXML))
	for _, f := range fzip.File {
		if f.Name == sheetXML {
			r, _ := f.Open()
			defer r.Close()

			buf := make([]byte, 1024*1024) // 1MB
			max := 0
			for {
				n, err := r.Read(buf)
				if n > 0 {
					v := maxRowFast(buf[:n])
					if v > max {
						max = v
					}
				}
				if err == io.EOF {
					break
				}
			}
			return max, nil
		}
	}
	return 0, fmt.Errorf("sheet not found")
}
func maxRowFast(data []byte) int {
	max := 0
	i := 0
	n := len(data)
	for i < n {
		// 找 r="
		if data[i] == 'r' && i+3 < n && data[i+1] == '=' && data[i+2] == '"' {
			j := i + 3
			v := 0
			for j < n && data[j] >= '0' && data[j] <= '9' {
				v = v*10 + int(data[j]-'0')
				j++
			}
			if v > max {
				max = v
			}
			i = j
		} else {
			i++
		}
	}
	return max
}
