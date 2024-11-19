package main

import (
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

func main() {
	inFolder := flag.String("in", "", "Folder containing files to merge")
	//outFile := flag.String("out", "merged.xlsx", "File name of output file")

	flag.Parse()
	if *inFolder == "" {
		var dir, err = os.Getwd()
		if err != nil {
			fmt.Println(err)
		}
		inFolder = &dir
	}

	out := excelize.NewFile()
	defer func() {
		if err := out.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	outSheets := out.GetSheetList()
	err := out.SetSheetName(outSheets[0], "Merged")
	if err != nil {
		fmt.Println(err)
	}

	format := "yyyy-mm-dd"
	dateStyle, err := out.NewStyle(&excelize.Style{CustomNumFmt: &format})
	if err != nil {
		fmt.Println(err)
	}
	headerStyle, _ := out.NewStyle(&excelize.Style{Font: &excelize.Font{Bold: true}})
	err = out.SetColStyle("Merged", "K:M", dateStyle)
	if err != nil {
		fmt.Println(err)
	}

	outRow := 2
	seen := make(map[string]bool)
	styles := make(map[string]int)
	entries, _ := os.ReadDir(*inFolder)
	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}

		name := entry.Name()
		ext := strings.ToLower(filepath.Ext(name))
		if name == "Merged.xlsx" || ext[:3] != ".xl" {
			continue
		}

		path := filepath.Join(*inFolder, name)
		fmt.Println(path)
		in, err := excelize.OpenFile(path)
		if err != nil {
			fmt.Println(err)
			return
		}
		defer func() {
			if err := in.Close(); err != nil {
				fmt.Println(err)
			}
		}()

		province := name
		pos := strings.Index(name, " ")
		if pos != -1 {
			province = name[:pos]
		}

		for _, sheet := range in.GetSheetList() {
			fmt.Println(sheet)
			if sheet[:3] != "202" {
				continue
			}
			dim, _ := in.GetSheetDimension(sheet)
			fmt.Println(dim)
			dim = dim[strings.Index(dim, ":")+1:]
			_, lastRow, _ := excelize.SplitCellName(dim)

			date := sheet
			pos := strings.Index(date, "_")
			if pos != -1 {
				date = date[:pos]
			}

			_ = out.SetColWidth("Merged", "A", "B", 15)
			_ = out.SetColWidth("Merged", "C", "C", 25)
			_ = out.SetColWidth("Merged", "D", "F", 15)
			_ = out.SetColWidth("Merged", "G", "J", 10)
			_ = out.SetColWidth("Merged", "K", "N", 15)

			_ = out.SetCellStr("Merged", "A1", "Province / Server")
			_ = out.SetCellStr("Merged", "B1", "Date")
			_ = out.SetCellStr("Merged", "C1", "Path")
			_ = out.SetCellStr("Merged", "D1", "Server")
			_ = out.SetCellStr("Merged", "E1", "Drive")
			_ = out.SetCellStr("Merged", "F1", "User")
			_ = out.SetCellStr("Merged", "G1", "Size")
			_ = out.SetCellStr("Merged", "H1", "Allocated")
			_ = out.SetCellStr("Merged", "I1", "Files")
			_ = out.SetCellStr("Merged", "J1", "Folders")
			_ = out.SetCellStr("Merged", "K1", "Creation Date")
			_ = out.SetCellStr("Merged", "L1", "Last Modified")
			_ = out.SetCellStr("Merged", "M1", "Last Accessed")
			_ = out.SetCellStr("Merged", "N1", "Owner")

			_ = out.SetRowStyle("Merged", 1, 1, headerStyle)

			for row := 7; row <= lastRow; row++ {
				value, _ := in.GetCellValue(sheet, "A"+strconv.Itoa(row))
				parts := removeEmpty(strings.Split(value, string(filepath.Separator)))
				key := date + filepath.Join(parts...)
				if seen[key] {
					continue
				}
				seen[key] = true

				if len(parts) == 3 {
					_ = out.SetCellStr("Merged", "A"+strconv.Itoa(outRow), province)
					_ = out.SetCellStr("Merged", "B"+strconv.Itoa(outRow), date)
					_ = out.SetCellStr("Merged", "C"+strconv.Itoa(outRow), value)
					_ = out.SetCellStr("Merged", "D"+strconv.Itoa(outRow), parts[0])
					_ = out.SetCellStr("Merged", "E"+strconv.Itoa(outRow), parts[1])
					_ = out.SetCellStr("Merged", "F"+strconv.Itoa(outRow), parts[2])

					addSize(in, out, sheet, "B"+strconv.Itoa(row), "G"+strconv.Itoa(outRow), styles)
					addSize(in, out, sheet, "C"+strconv.Itoa(row), "H"+strconv.Itoa(outRow), styles)

					addNumber(in, out, sheet, "D"+strconv.Itoa(row), "I"+strconv.Itoa(outRow))
					addNumber(in, out, sheet, "E"+strconv.Itoa(row), "J"+strconv.Itoa(outRow))

					addDate(in, out, sheet, "F"+strconv.Itoa(row), "K"+strconv.Itoa(outRow))
					addDate(in, out, sheet, "G"+strconv.Itoa(row), "L"+strconv.Itoa(outRow))
					addDate(in, out, sheet, "H"+strconv.Itoa(row), "M"+strconv.Itoa(outRow))

					value, _ := in.GetCellValue(sheet, "I"+strconv.Itoa(row))
					_ = out.SetCellStr("Merged", "N"+strconv.Itoa(outRow), value)

					outRow++
				}
			}
		}
	}

	err = out.SetPanes("Merged", &excelize.Panes{
		Freeze:      true,
		Split:       false,
		XSplit:      0,
		YSplit:      1,
		TopLeftCell: "A2",
		ActivePane:  "topRight",
	})
	if err != nil {
		fmt.Println(err)
	}

	err = out.AutoFilter("Merged", "A1:N"+strconv.Itoa(outRow-1), []excelize.AutoFilterOptions{})
	if err != nil {
		fmt.Println(err)
	}

	err = out.SaveAs(filepath.Join(*inFolder, "Merged.xlsx"))
	if err != nil {
		fmt.Println(err)
	}
}

func addSize(in *excelize.File, out *excelize.File, sheet string, inCell string, outCell string, styles map[string]int) {
	styleId, err := in.GetCellStyle(sheet, inCell)
	if err != nil {
		fmt.Println("STYLE ERROR", err)
	}
	var outStyleId = 0
	var ok bool
	if styleId > 0 {
		style, _ := in.GetStyle(styleId)
		format := *style.CustomNumFmt
		if strings.Contains(format, "MB") {
			format = strings.ReplaceAll(format, "MB", "GB")
		} else if strings.Contains(format, "KB") {
			format = strings.ReplaceAll(format, "KB", "GB")
		} else if strings.Contains(format, "Bytes") {
			format = strings.ReplaceAll(format, "Bytes", "GB")
		}
		outStyleId, ok = styles[format]
		if !ok {
			style.CustomNumFmt = &format
			outStyleId, err = out.NewStyle(style)
			if err != nil {
				fmt.Println(err)
			}
			styles[format] = outStyleId
		}
	}

	divider := 1
	style, _ := in.GetStyle(styleId)
	format := *style.CustomNumFmt
	if strings.Contains(format, "MB") {
		divider = 1024
	} else if strings.Contains(format, "KB") {
		divider = 1024 * 1024
	} else if strings.Contains(format, "Bytes") {
		divider = 1024 * 1024 * 1024
	}

	inValue, _ := in.GetCellValue(sheet, inCell, excelize.Options{RawCellValue: true})
	floatValue, _ := strconv.ParseFloat(inValue, 32)
	floatValue /= float64(divider)
	_ = out.SetCellStyle("Merged", outCell, outCell, outStyleId)
	_ = out.SetCellValue("Merged", outCell, floatValue)
}

func addNumber(in *excelize.File, out *excelize.File, sheet string, inCell string, outCell string) {
	inValue, _ := in.GetCellValue(sheet, inCell, excelize.Options{RawCellValue: true})
	intValue, _ := strconv.ParseInt(inValue, 10, 64)
	_ = out.SetCellValue("Merged", outCell, intValue)
}

func addDate(in *excelize.File, out *excelize.File, sheet string, inCell string, outCell string) {
	inValue, _ := in.GetCellValue(sheet, inCell, excelize.Options{RawCellValue: true})
	floatValue, _ := strconv.ParseFloat(inValue, 64)
	_ = out.SetCellValue("Merged", outCell, floatValue)
}

func removeEmpty(in []string) []string {
	var r []string
	for _, str := range in {
		if str != "" {
			r = append(r, str)
		}
	}
	return r
}
